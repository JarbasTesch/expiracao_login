from collections import defaultdict
import pandas as pd
import funcoes as fcs

df = pd.read_excel('login_expirar_teste.xlsx')
colaboradores = defaultdict(lambda:{"nome_colaborador": None, "gestores":set(), "data_expiracao": None, "dias_restantes": None, "login": None, "UO": None})

for _, row in df.iterrows():
    nome_colaborador = row['NOME DO USUARIO']
    email_colaborador = row['EMAIL DO USUARIO']
    gestor = row['EMAIL DO GESTOR']
    data_expiracao = row['DATA DA EXPIRAÇÃO']
    dias_restantes = row['DIAS A EXPIRAR']
    login = row['LOGIN']
    uo = row['UO']

    colaboradores[email_colaborador]["email_colaborador"] = email_colaborador
    colaboradores[email_colaborador]["nome_colaborador"] = nome_colaborador
    colaboradores[email_colaborador]["gestores"].add(gestor)
    colaboradores[email_colaborador]["data_expiracao"] = data_expiracao
    colaboradores[email_colaborador]["dias_restantes"] = dias_restantes
    colaboradores[email_colaborador]["login"] = login
    colaboradores[email_colaborador]["UO"] = uo

for email_colaborador, info in colaboradores.items():
    nome = info["nome_colaborador"]
    gestores = list(info["gestores"])
    data_expiracao = info["data_expiracao"]
    dias_restantes = info["dias_restantes"]
    login = info["login"]
    uo = info["UO"]

    fcs.enviar_email(email_colaborador, nome, gestores, data_expiracao, dias_restantes, login, uo)