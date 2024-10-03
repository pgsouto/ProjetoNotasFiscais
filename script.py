import imaplib
import email
from email.header import decode_header
import os
from pdf2image import convert_from_path
import pytesseract
import socket
import re
from datetime import datetime
import smtplib
from email.mime.text import MIMEText

# Configuração de conexão com o e-mail
IMAP_SERVER = "imap.gmail.com"
EMAIL_ACCOUNT = "jmf.boys27@gmail.com"
PASSWORD = "hnyg sjdt jwwj nxrd"

# Defina o timeout para 10 segundos
socket.setdefaulttimeout(10)

# Conectar ao servidor IMAP
print("Conectando ao servidor...")
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL_ACCOUNT, PASSWORD)
print("Logado com sucesso.")

print("Selecionando a caixa de entrada...")
mail.select('inbox')

# Buscar os e-mails mais recentes (últimos 10)
status, emails = mail.search(None, f'SUBJECT "TeclaT: Nota Fiscal"')
email_ids = emails[0].split()

# Limitar a busca aos últimos 10 e-mails, se houver
if len(email_ids) > 10:
    email_ids = email_ids[-10:]

# Ler e-mails e baixar os anexos PDF
for num in email_ids:
    status, data = mail.fetch(num, '(RFC822)')
    raw_email = data[0][1]
    email_message = email.message_from_bytes(raw_email)

    # Percorrer os anexos
    for part in email_message.walk():
        if part.get_content_type() == "application/pdf":
            file_name = part.get_filename()

            # Decodificar o nome do arquivo se estiver em formato MIME
            if file_name:
                decoded_name, encoding = decode_header(file_name)[0]
                if isinstance(decoded_name, bytes):
                    # Decodificar bytes para string, com o encoding correto
                    file_name = decoded_name.decode(encoding if encoding else 'utf-8')

                # Corrigir caracteres inválidos no nome do arquivo
                file_name = file_name.replace("/", "").replace("\\", "")

                print(f"Baixando arquivo: {file_name}")
                with open(file_name, 'wb') as f:
                    f.write(part.get_payload(decode=True))

                # Converter PDF para imagem e usar OCR para extrair o texto
                print(f"Convertendo PDF {file_name} em imagem e extraindo texto...")
                pages = convert_from_path(file_name, 500)
                full_text = ""
                for page in pages:
                    config = '--psm 0'
                    text = pytesseract.image_to_string(page)
                    full_text += text

                # Exibir o texto extraído
                print("Texto extraído da nota fiscal:")
                print(full_text)

agora = datetime.now()

meses = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro"
}

def encontrar_data_competencia(texto):
    # Padrão de data no formato dd/mm/aaaa
    padrao_data = r'(\d{2}/\d{2}/\d{4})'
    datas_encontradas = re.findall(padrao_data, texto)

    if datas_encontradas:
        return datas_encontradas[0]  # Retorna a primeira data
    return None

#modificar a lógica posteriormente para a RN 
def verificar_data_competencia(data):
    mes_anterior = agora.month
    ano_atual = agora.year
    dia_competencia, mes_competencia, ano_competencia = map(int, data.split("/"))
    if(mes_competencia == mes_anterior and ano_competencia == ano_atual):
        return True
    else:
        return False
    
def encontrar_descricao_servico(texto):
    # Regex para capturar a descrição do serviço
    padrao = r"(?:Descrigao do Servicgo|Servicgo Prestado|Tipo de Servicgo|Objeto)[\s:]*([\w\s,.:-]+?)(?=\.\s|\n|$)"
    resultado = re.search(padrao, texto, re.IGNORECASE)
    
    if resultado:
        return resultado.group(1).strip()  
    return None 

#modificar a lógica posteriormente para a RN 
def verificar_data_descricao(descricao):
    mes_anterior = agora.month-1
    padrao_meses = r'|'.join(re.escape(mes) for mes in meses.values())
    resultado = re.findall(padrao_meses, full_text, re.IGNORECASE)

    mes_encontrado = None
    if resultado:
        mes_nome = resultado[0]  # Pega o primeiro mês encontrado
        # Encontra o número do mês no dicionário
        mes_encontrado = next((num for num, nome in meses.items() if nome.lower() == mes_nome.lower()), None)
        if mes_encontrado == mes_anterior:
            return True
            print("a descrição de serviço está correta. ✓\n")
        else:
            return False
            print("\033 a descrição de serviço está incorreta \033\n")

    #print(f"Número do mês encontrado: {mes_encontrado}")

def encontrar_valor_servico(texto):
    padrao = r"Valor do Servigo Desconto Incondicionado\s*R\$\s*([\d.,]+)"
    resultado = re.search(padrao, texto, re.IGNORECASE)

    if resultado:
        return float(resultado.group(1).replace('.', '').replace(',', '.')) # Retorna o valor capturado
    else:
        return None
    
def verificar_issqn(texto):
    # Usando regex para encontrar "Nao Retido" (considerando que pode haver espaços)
    padrao = r"(?i)Retengao do ISSQN\s+([^-\n]*)"
    resultado = re.search(padrao, texto)

    if resultado:
        issqn_status = resultado.group(1).strip()
        if "Nao Retido" in issqn_status:
            return True
    return False

def verificar_valor_servico(valor_servico, texto):

    valor_liquido = encontrar_valor_liquido(texto)
    if verificar_issqn(texto) and valor_liquido is not None and valor_servico == valor_liquido and verificar_issqn(texto):
        return True
    else:
        return False

def encontrar_valor_liquido(texto):
    # Padrão para encontrar o valor líquido
    # Esse padrão assume que o valor líquido aparece após uma expressão como "Valor Líquido" ou similar
    padrao = r"Valor Liquido da NFS-e\s*R\$\s*([\d.,]+)"
    resultado = re.search(padrao,text, re.IGNORECASE)

    if resultado:
        return float(resultado.group(1).replace('.', '').replace(',', '.'))  # Converte para float
    return None  # Retorna None se o valor não for encontrado
    
data_competencia = encontrar_data_competencia(full_text)
if data_competencia:
    print(f"Data de competência encontrada: {data_competencia}")
    if verificar_data_competencia(data_competencia):
        print("A competência está correta. ✓\n")
    else:
        print("A competência está incorreta.\n")


else:
    print("Nenhuma data de competência encontrada.")

descricao_servico = encontrar_descricao_servico(full_text)
if descricao_servico:
    print(f"Descrição de serviço encontrada: {descricao_servico}")
    if verificar_data_descricao(descricao_servico):
        print("A descrição de serviço está correta. ✓\n")
    else:
         print("A descrição de serviço está incorreta.\n")

else:
    print("Nenhuma descrição de serviço encontrada.")

verificar_data_descricao(descricao_servico)

valor_servico = encontrar_valor_servico(full_text)
if valor_servico:
    print(f"Valor do Serviço encontrado: R$ {valor_servico}")
else:
    print("Valor do Serviço não encontrado.")

valor_liquido = encontrar_valor_liquido(full_text)
if valor_liquido is not None:
    print(f"Valor Líquido encontrado: R$ {valor_liquido:.2f}\n")
else:
    print("Valor Líquido não encontrado.\n")


if verificar_issqn(full_text):
    print("A retenção ISSQN consta como: Não Retido.")
else:
    print("Retenção ISSQN incorreta.")

if verificar_valor_servico(valor_servico, full_text):
    print("O valor do serviço está correto. ✓")
else:
    print("\033O valor do serviço está incorreto.\033")

# Função para enviar e-mail
def enviar_email(subject, body, to_email):
    # Configurações do servidor SMTP
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    smtp_username = "jmf.boys27@gmail.com"
    smtp_password = "hnyg sjdt jwwj nxrd"

    # Criação do objeto MIME
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = smtp_username
    msg["To"] = "pedro.guilherme@teclat.com.br"

    try:
        # Conexão e envio do e-mail
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Ativa a criptografia TLS
            server.login(smtp_username, smtp_password)
            server.sendmail(smtp_username, to_email, msg.as_string())
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Falha ao enviar e-mail: {e}")

# Verificações
competencia_correta = verificar_data_competencia(data_competencia)
descricao_correta = verificar_data_descricao(descricao_servico)  # Verifique se a descrição foi encontrada
valor_servico_correto = verificar_valor_servico(valor_servico, full_text)

# Compor o corpo do e-mail baseado nas verificações
if competencia_correta and descricao_correta and valor_servico_correto:
    subject = "Resultado da Verificação da Nota Fiscal"
    body = "Tudo está correto com a nota fiscal."
else:
    subject = "Resultado da Verificação da Nota Fiscal"
    body = "A nota fiscal está incorreta. As seguintes verificações falharam:\n"
    
    if not competencia_correta:
        body += "- Competência incorreta.\n"
    if not descricao_correta:
        body += "- Descrição do serviço incorreta.\n"
    if not valor_servico_correto:
        body += "- Valor do serviço incorreto.\n"

# Enviar e-mail com o resultado
enviar_email(subject, body, "pedro.guilherme@teclat.com.br")


# Encerrar a conexão com o servidor de e-mail
mail.logout()
print("Conexão encerrada.")