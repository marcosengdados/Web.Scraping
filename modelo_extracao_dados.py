import pandas as pd
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import time

def abrir_links_e_extrair_info_excel(caminho_arquivo, nome_coluna_link):

    # Carrega os dados
    df = pd.read_excel(caminho_arquivo)
    resultados = []

    for index, linha in df.iterrows():
        url = linha.get(nome_coluna_link)
        nome_empreendimento = linha.get("EMPREENDIMENTO x")

        if pd.isna(url) or not str(url).startswith("http"):
            print(f"URL inválida na linha {index}: {url}")
            continue

        try:
            response = requests.get(url)
            soup = BeautifulSoup(response.content, "html.parser")

            # Busca os andamentos do processo
            tabela_andamentos = soup.find_all("tr", class_=["andamentox", "andamentoConcluido"])
            datas = []

            for tr in tabela_andamentos:
                tds = tr.find_all("td")
                if len(tds) >= 1:
                    try:
                        data_str = tds[0].text.strip()
                        data = datetime.strptime(data_str, "%d/%m/%Y %H:%M")
                        datas.append((data, tds[1].text.strip(), tds[2].text.strip()))
                    except:
                        pass

            if datas:
                mais_recente = max(datas, key=lambda x: x[0])
                data_formatada = mais_recente[0].strftime("%d/%m/%Y %H:%M")
                unidade = mais_recente[1]
                descricao = mais_recente[2]
            else:
                data_formatada = unidade = descricao = "Não encontrado"

            # Busca o interessado
            interessados_td = soup.find("td", string="Interessados:")
            if interessados_td:
                interessado_texto = interessados_td.find_next_sibling("td").text.strip()
            else:
                interessado_texto = "Não encontrado"

            # Salva os dados
            resultados.append({
                "Link": url,
                "Nome do Empreendimento": nome_empreendimento,
                "Atualização mais recente": data_formatada,
                "Unidade": unidade,
                "Descrição": descricao,
                "Interessado": interessado_texto
            })

            time.sleep(3)  # Evita sobrecarregar o servidor

        except Exception as e:
            print(f"Erro ao acessar {url}: {e}")

    # Exporta o resultado para Excel
    df_resultado = pd.DataFrame(resultados)
    df_resultado.to_excel("Resumo_extraido.xlsx", index=False)
    print("Resumo extraído com sucesso.")
    print(df_resultado)

# Executa a função
abrir_links_e_extrair_info_excel("x.xlsx", "LINK DE ACESSO x")

import pandas as pd
import matplotlib.pyplot as plt
import win32com.client as win32
import os

# 1. Leitura do arquivo Excel
processos_sei = pd.read_excel('Resumo_extraido.xlsx')
pd.set_option('display.max_columns', None)

# 2. Limpa espaços dos nomes de colunas
processos_sei.columns = processos_sei.columns.str.strip()

# 3. Seleção das colunas
caracteristica = processos_sei[['x Empreendimento', 'Atualização mais recente', 'Descrição']]

# 4. Caminho do arquivo Excel e da imagem da logo
arquivo = r'C:\Users\Admin\PycharmProjects\pythonProject2\x.xlsx'
logo_path = r'C:\Users\Admin\PycharmProjects\pythonProject2\x.png'
logo_cid = 'minhaLogo'  # Identificador do CID da imagem

corpo_email = f'''
<p>Prezados(as) Analistas,</p>

<p>Encaminho, abaixo, o Relatório da Lista de Acompanhamento referente a x Engenharia.</p>

<p>O documento foi elaborado com base em dados atualizados, organizados de forma a facilitar a visualização das informações relevantes, como o nome dos x, a data da última atualização e as respectivas descrições.</p>

<p>Ressalto que o relatório busca oferecer maior clareza no acompanhamento técnico e administrativo das demandas em andamento.</p>

<p>Qualquer dúvida, estou à disposição.</p>

<p>
<p>Atenciosamente,<br>
Fulano da Silva<br>
x Engenharia</p>

<div style="text-align:center; margin-top:20px;">
    <img src="cid:{logo_cid}" style="width:450px; opacity:0.9;" alt="x">
</div>

<p style="font-style: italic; font-size: 11px; text-align: center; color: #555; margin-top: 10px;">
    O conteúdo deste e-mail é confidencial e destinado exclusivamente ao destinatário especificado apenas na mensagem. 
    É estritamente proibido compartilhar qualquer parte desta mensagem com terceiros, sem o consentimento por escrito do remetente. <br>
    (The content of this email is confidential and intended exclusively for the recipient specified in the message only. 
    It is strictly prohibited to share any part of this message with third parties without the written consent of the sender.)
</p>
'''

# 6. Envio do e-mail via Outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'x@gmail.com'
mail.Subject = 'Relatório da Lista de Acompanhamento'
mail.HTMLBody = corpo_email

# 7. Anexa o Excel, se encontrado
if os.path.exists(arquivo):
    mail.Attachments.Add(arquivo)
else:
    print('⚠️ Arquivo Excel não encontrado!')

# 8. Anexa a imagem como embedded (com CID), se encontrada
if os.path.exists(logo_path):
    attachment = mail.Attachments.Add(logo_path)
    # Define o CID para que a imagem seja incorporada ao corpo do e-mail
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", logo_cid)
else:
    print('⚠️ Imagem da logo não encontrada!')

# 9. Envia o e-mail
mail.Send()
print('✅ E-mail enviado com sucesso pelo Outlook.')