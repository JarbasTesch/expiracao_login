import win32com.client as win32
from datetime import datetime

def enviar_email(email_colaborador, nome_colaborador, gestor, data_expiracao, dias_restantes, login, uo):

    destinatarios = f"{email_colaborador}; {'; '.join(gestor)}"
    data_formatada = data_expiracao.strftime("%d/%m/%Y")

    outlook = win32.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    mail.To = destinatarios
    #mail.SentOnBehalfOfName = 'pdv2024@eletronuclear.gov.br'
    mail.Subject = 'Expiração de login'
    mail.HTMLBody =f"""
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Expiração de Login</title>
            <style>
                table {{
                    border-collapse: collapse;
                    width: 100%;
                }}
                th, td {{
                    border: 1px solid black;
                    padding: 8px;
                    text-align: center; /* Alinha o texto ao centro */
                }}
                th {{
                    background-color: #f2f2f2;
                }}
            </style>
        </head>
        <body>
            <p><strong>Prezado(a) {nome_colaborador},</strong></p>
            <p>Informamos que a <strong>Expiração de Login</strong> está próxima.</p>
            <p>Sua conta irá expirar no dia {data_formatada}, daqui a {dias_restantes} dias.</p>
            <p>Caso precise de ajuda, entre em contato com seus gestores: {', '.join(gestor)}</strong></p>
        
            <h3>Detalhes da Conta:</h3>
            <table>
                <thead>
                    <tr>
                        <th>Nome</th>
                        <th>Login</th>
                        <th>UO</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>{nome_colaborador}</td>
                        <td>{login}</td>
                        <td>{uo}</td>
                    </tr>
                </tbody>
            </table>
        
            <p>Atenciosamente,</p>
            <p>Equipe de TI</p>
        </body>
        </html>
        """
    mail.Send()
    print(f'Email enviado para: {destinatarios}')