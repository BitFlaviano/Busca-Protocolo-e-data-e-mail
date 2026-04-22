import imaplib
import email
import re
from bs4 import BeautifulSoup
from email.utils import parsedate_to_datetime

IMAP_SERVER = "imap.titan.email"
EMAIL = "seuemail@provedor"
SENHA = "pass"


def limpar_html(html):
    try:
        soup = BeautifulSoup(html, "html.parser")
        return soup.get_text(separator=" ")
    except:
        return html


def extrair_texto(msg):
    corpo = ""

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            payload = part.get_payload(decode=True)

            if payload:
                try:
                    texto = payload.decode(errors="ignore")

                    if content_type == "text/plain":
                        corpo += texto
                    elif content_type == "text/html":
                        corpo += limpar_html(texto)

                except:
                    continue
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            try:
                corpo = payload.decode(errors="ignore")
            except:
                pass

    return corpo


def extrair_data(msg):
    try:
        data_raw = msg.get("Date")
        if data_raw:
            dt = parsedate_to_datetime(data_raw)
            return dt.strftime("%d/%m/%Y %H:%M")
    except:
        pass

    return "Data desconhecida"


def extrair_protocolos(texto):
    protocolos = []

    padrao1 = re.findall(r'protocolo.*?:\s*(\d+)', texto, re.IGNORECASE)
    padrao2 = re.findall(r'protocolo.*?(\d{8,20})', texto, re.IGNORECASE)

    protocolos.extend(padrao1)
    protocolos.extend(padrao2)

    return list(set(protocolos))


def buscar_protocolos():
    resultados = []

    try:
        print("Conectando...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)

        print("Logando...")
        mail.login(EMAIL, SENHA)

        mail.select("inbox")

        status, mensagens = mail.search(None, '(OR FROM "vivo" SUBJECT "vivo")')

        lista_ids = mensagens[0].split()
        print(f"Total de emails encontrados: {len(lista_ids)}")

        for i, num in enumerate(lista_ids, 1):
            try:
                status, data = mail.fetch(num, "(RFC822)")
                msg = email.message_from_bytes(data[0][1])

                texto = extrair_texto(msg)
                data_email = extrair_data(msg)

                protocolos = extrair_protocolos(texto)

                for p in protocolos:
                    resultados.append({
                        "protocolo": p,
                        "data": data_email
                    })

                if protocolos:
                    print(f"[{i}] {data_email} → {protocolos}")

            except Exception as e:
                print(f"Erro no email {num}: {e}")

        mail.logout()

    except Exception as e:
        print("Erro geral:", e)

    return resultados


if __name__ == "__main__":
    lista = buscar_protocolos()

    print("\n==== RESULTADO FINAL ====")
    for item in lista:
        print(item["data"], "-", item["protocolo"])

    # salvar Excel
    if lista:
        try:
            import pandas as pd
            df = pd.DataFrame(lista)
            df.to_excel("protocolos.xlsx", index=False)
            print("\nArquivo protocolos.xlsx criado!")
        except:
            print("\nInstale: pip install pandas openpyxl")