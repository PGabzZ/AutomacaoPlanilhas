import os
from datetime import datetime
import pandas as pd
import win32com.client as win32

caminho = "Bases"
arquivos = os.listdir(caminho)
tabela_consolidada = pd.DataFrame()

for planilha in arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho, planilha))
    tabela_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas["Data de Venda"],
                                                                                     unit = "d")
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda")
tabela_consolidada = tabela_consolidada.reset_index(drop=True)
tabela_consolidada.to_excel("Vendas.xlsx", index=False)

outlook = win32.Dispatch("outlook.application")
email = outlook.CreateItem(0)
email.to = ""
currentdate = datetime.today().strftime("%d/%m/%Y")
email.Subject = f"Relatório de vendas {currentdate}"
email.Body= f"""
Prezado,

Segue em anexo relatório de vendas do dia {currentdate}.
Qualquer dúvida estou a disposição
"""

caminhoemail = os.getcwd()
attachment = os.path.join(caminhoemail, "Vendas.xlsx")
email.Attachments.Add(attachment)

email.Send()



