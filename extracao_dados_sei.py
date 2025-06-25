import os
import threading
import time
from datetime import datetime

import pandas as pd
import requests
from bs4 import BeautifulSoup
from flask import Flask, render_template, redirect, url_for
import win32com.client as win32
from ftfy import fix_text
app = Flask(__name__)

CAMINHO_EXCEL = "LISTA_DE_ACOMPANHAMENTO.xlsx"
COLUNA_URL = "LINK DE ACESSO PÚBLICO"
EMAIL_DESTINO = "x.com"
ARQUIVO_SAIDA = "Resumo_extraido.xlsx"

ultima_verificacao = None
ultima_data_enviada = None

with open("api.txt", "r") as file:
    SCRAPERAPI_KEY = file.read().strip()

def pode_executar_agora():
    agora = datetime.now()
    dia_semana = agora.weekday()
    hora_atual = agora.time()
    inicio = datetime.strptime("08:00", "%H:%M").time()
    fim = datetime.strptime("17:00", "%H:%M").time()
    feriados = ["01-01", "04-21", "05-01", "09-07", "10-12", "11-02", "11-15", "12-25"]
    hoje = agora.strftime("%m-%d")
    return dia_semana < 5 and inicio <= hora_atual <= fim and hoje not in feriados

def abrir_links_e_extrair_info_excel(caminho_arquivo, nome_coluna_link):
    df = pd.read_excel(caminho_arquivo)
    resultados = []

    for index, linha in df.iterrows():
        url_original = linha.get(nome_coluna_link)
        nome_empreendimento = linha.get("EMPREENDIMENTO")

        if pd.isna(url_original) or not str(url_original).startswith("http"):
            continue

        scraperapi_url = "http://api.scraperapi.com/"
        payload = {
            "api_key": SCRAPERAPI_KEY,
            "url": url_original
        }

        try:
            response = requests.get(scraperapi_url, params=payload, timeout=60)
            html_decodificado = response.content.decode("ISO-8859-1", errors="replace")
            soup = BeautifulSoup(html_decodificado, "html.parser")

            tabela_andamentos = soup.find_all("tr", class_=["andamentoAberto", "andamentoConcluido"])
            datas = []

            for tr in tabela_andamentos:
                tds = tr.find_all("td")
                if len(tds) >= 3:
                    try:
                        data_str = tds[0].get_text(strip=True)
                        unidade = tds[1].get_text(strip=True)
                        descricao_bruta = tds[2].get_text(strip=True)

                        # ✅ Correção definitiva usando ftfy
                        descricao = fix_text(descricao_bruta)

                        data = datetime.strptime(data_str, "%d/%m/%Y %H:%M")
                        datas.append((data, unidade, descricao))
                    except Exception as e:
                        print(f"Erro ao processar linha da tabela: {e}")

            if datas:
                mais_recente = max(datas, key=lambda x: x[0])
                data_formatada = mais_recente[0].strftime("%d/%m/%Y %H:%M")
                unidade = mais_recente[1]
                descricao = mais_recente[2]
            else:
                data_formatada = unidade = descricao = "Não encontrado"

            interessados_td = soup.find("td", string="Interessados:")
            interessado_texto = interessados_td.find_next_sibling("td").text.strip() if interessados_td else "Não encontrado"

            resultados.append({
                "Link": url_original,
                "Nome do Empreendimento": nome_empreendimento,
                "Atualização mais recente": data_formatada,
                "Unidade": unidade,
                "Descrição": descricao,
                "Interessado": interessado_texto
            })

            time.sleep(3)

        except Exception as e:
            print(f"Erro em {url_original}: {e}")

    arquivo_saida = os.path.abspath("Resumo_extraido.xlsx")
    pd.DataFrame(resultados).to_excel(arquivo_saida, index=False)
    return arquivo_saida

def enviar_email_se_atualizacao(arquivo_excel, email_destino):
    if not os.path.exists(arquivo_excel):
        return

    processos = pd.read_excel(arquivo_excel)
    processos.columns = processos.columns.str.strip()
    processos['Atualização mais recente'] = pd.to_datetime(
        processos['Atualização mais recente'], format='%d/%m/%Y %H:%M', errors='coerce'
    )

    data_hoje = datetime.now().date()
    atualizacao_hoje = processos['Atualização mais recente'].dt.date == data_hoje

    if atualizacao_hoje.any():
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = email_destino
        mail.Subject = 'Relatório de Atualizações SEI'
        mail.HTMLBody = """
        <p>Prezados,</p>
        <p>Segue relatório atualizado com os andamentos de hoje.</p>
        <p>Atenciosamente,<br>xx Engenharia</p>
        """
        mail.Attachments.Add(arquivo_excel)
        mail.Send()

def scraping_loop():
    global ultima_verificacao, ultima_data_enviada

    while True:
        agora = datetime.now()
        ultima_verificacao = agora.strftime("%d/%m/%Y %H:%M:%S")

        if pode_executar_agora():
            try:
                print("🔄 Rodando verificação automática...")
                arquivo_gerado = abrir_links_e_extrair_info_excel(CAMINHO_EXCEL, COLUNA_URL)
                df = pd.read_excel(arquivo_gerado)
                df['Atualizacao'] = pd.to_datetime(df['Atualização mais recente'], format='%d/%m/%Y %H:%M', errors='coerce')

                novas = df[df['Atualizacao'] > (ultima_data_enviada or datetime.min)]

                if not novas.empty:
                    enviar_email_se_atualizacao(arquivo_gerado, EMAIL_DESTINO)
                    ultima_data_enviada = df['Atualizacao'].max()
                    print("📬 E-mail enviado com novas atualizações.")
                else:
                    print("📭 Nenhuma atualização desde o último envio.")

            except Exception as e:
                print(f"❌ Erro: {e}")
        else:
            print("🕓 Fora do horário permitido. Aguardando...")

        time.sleep(1800)

@app.route("/")
def index():
    dados = []
    if os.path.exists(ARQUIVO_SAIDA):
        df = pd.read_excel(ARQUIVO_SAIDA)
        dados = df.to_dict(orient="records")

    return render_template("index.html",
                           ultima_verificacao=ultima_verificacao,
                           ultima_data_enviada=ultima_data_enviada.strftime("%d/%m/%Y %H:%M:%S") if ultima_data_enviada else "Nunca",
                           dados=dados)

@app.route("/forcar")
def forcar_verificacao():
    threading.Thread(target=forcar_execucao).start()
    return redirect(url_for("index"))

def forcar_execucao():
    global ultima_data_enviada, ultima_verificacao
    agora = datetime.now()
    ultima_verificacao = agora.strftime("%d/%m/%Y %H:%M:%S")
    print("⚙️ Execução manual iniciada...")
    try:
        arquivo_gerado = abrir_links_e_extrair_info_excel(CAMINHO_EXCEL, COLUNA_URL)
        df = pd.read_excel(arquivo_gerado)
        df['Atualizacao'] = pd.to_datetime(df['Atualização mais recente'], format='%d/%m/%Y %H:%M', errors='coerce')
        novas = df[df['Atualizacao'] > (ultima_data_enviada or datetime.min)]
        if not novas.empty:
            enviar_email_se_atualizacao(arquivo_gerado, EMAIL_DESTINO)
            ultima_data_enviada = df['Atualizacao'].max()
            print("📨 E-mail enviado por execução manual.")
    except Exception as e:
        print(f"❌ Erro manual: {e}")

if __name__ == "__main__":
    threading.Thread(target=scraping_loop, daemon=True).start()
    app.run(debug=True)






