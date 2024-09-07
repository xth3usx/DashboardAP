import os.path
import subprocess
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import matplotlib.pyplot as plt
import numpy as np  

# Escopo para acessar a planilha do Google Sheets
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# ID da planilha. A faixa de células será dinâmica
SAMPLE_SPREADSHEET_ID = "14IkU4ZDRoDYraeV4DiPGShlnbCUYOszJiMR4eqHav8M"

def ping_ip(ip):
    try:
        # Executa o comando ping
        output = subprocess.check_output(["ping", "-n", "1", ip], universal_newlines=True)
        # Verifica se "TTL" está na saída do comando, o que indica sucesso
        if "TTL" in output:
            return True
        else:
            return False
    except subprocess.CalledProcessError:
        return False

def generate_donut_chart(available_count, unavailable_count):
    # Definir os dados para o gráfico de rosca
    labels = ['Disponíveis', 'Indisponíveis']
    sizes = [available_count, unavailable_count]
    colors = ['#4CAF50', '#F44336']  # Verde para disponíveis, vermelho para indisponíveis

    # Criar o gráfico de rosca mais estilizado
    fig, ax = plt.subplots(figsize=(5, 5), subplot_kw=dict(aspect="equal"))

    wedges, texts, autotexts = ax.pie(
        sizes, 
        colors=colors, 
        startangle=90, 
        counterclock=False, 
        wedgeprops=dict(width=0.5), 
        autopct='%1.1f%%',  # Mostrando porcentagens
        textprops=dict(color="black", fontsize=10)
    )

    # Centralizar o título no gráfico
    plt.setp(autotexts, size=10, weight="bold")
    ax.set_title("Status dos Dispositivos", fontsize=16, fontweight='bold')

    # Melhorar a legenda
    ax.legend(wedges, labels, title="Status", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

    # Salvar o gráfico
    plt.savefig("donut_chart.png", bbox_inches='tight', transparent=True)
    plt.close()

def generate_html_table(data, available_count, unavailable_count):
    # Data e hora da última atualização
    last_updated = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    # Gerar o gráfico de rosca
    generate_donut_chart(available_count, unavailable_count)

    # HTML com cabeçalho, menu e estilo melhorado
    html = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background-color: #f4f7f6;
                margin: 0;
                padding: 20px;
            }}
            .header {{
                text-align: center;
                padding: 10px 0;
                margin-bottom: 10px;
                border-bottom: 2px solid #ddd;
            }}
            .header h1 {{
                font-size: 32px;
                color: #333;
                margin: 0;
                font-weight: 600;
            }}
            .updated-time {{
                font-size: 14px;
                color: #777;
                margin-top: 10px;
            }}
            .status-boxes {{
                display: flex;
                justify-content: center;
                gap: 15px;
                margin: 10px 0;
                font-size: 16px;
                color: white;
            }}
            .status-box {{
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: 600;
            }}
            .available-box {{
                background-color: #4CAF50;  /* Verde elegante */
            }}
            .unavailable-box {{
                background-color: #F44336;  /* Vermelho elegante */
            }}
            .menu {{
                text-align: center;
                margin: 20px 0;
                padding: 10px 0;
                border-top: 2px solid #ddd;
                border-bottom: 2px solid #ddd;
            }}
            .menu a {{
                text-decoration: none;
                background-color: #0078d4;
                color: white;
                padding: 8px 16px;
                border-radius: 5px;
                font-size: 14px;
                margin: 0 10px;
                transition: background-color 0.3s;
            }}
            .menu a:hover {{
                background-color: #005bb5;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                background-color: #ffffff;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 0 15px rgba(0, 0, 0, 0.1);
                display: flex;
                flex-wrap: wrap;
                gap: 10px;
            }}
            .label-box {{
                padding: 10px;
                border-radius: 50%;
                box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.1);
                cursor: pointer;
                transition: transform 0.2s, box-shadow 0.2s;
                width: 50px;
                height: 50px;
                text-align: center;
                font-size: 14px;
                font-weight: bold;
                text-decoration: none;
            }}
            .label-box:hover {{
                transform: scale(1.05);
                box-shadow: 0px 0px 20px rgba(0, 0, 0, 0.3);
            }}
            .available {{
                background-color: #C8E6C9;
                color: #2E7D32;
            }}
            .unavailable {{
                background-color: #FFCDD2;
                color: #C62828;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>Dashboard de Conectividade dos Dispositivos</h1>
            <div class="updated-time">Última atualização: {last_updated}</div>
            <div class="status-boxes">
                <div class="status-box available-box">Disponíveis: {available_count}</div>
                <div class="status-box unavailable-box">Indisponíveis: {unavailable_count}</div>
            </div>
        </div>
        <div class="menu">
            <a href="https://example1.com">Item 1</a>
            <a href="https://example2.com">Item 2</a>
            <a href="https://example3.com">Item 3</a>
            <a href="https://example4.com">Item 4</a>
            <a href="https://example5.com">Item 5</a>
        </div>
        <div class="container">
    """

    # Iterar sobre os dados e construir as caixas com links para os IPs
    for row in data[1:]:
        ip = row[0] if len(row) > 0 else "N/A"
        etiqueta = row[1] if len(row) > 1 else "N/A"
        status_class = "available" if ping_ip(ip) else "unavailable"
        html += f"<a href='http://{ip}' target='_blank' class='label-box {status_class}'>{etiqueta}</a>"

    # Fechar o HTML
    html += """
        </div>
    </body>
    </html>
    """

    return html

def main():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=8080)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Consultar todos os dados da planilha, sem limitar o número de linhas
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Dashboard").execute()
        data = result.get("values", [])

        if not data:
            print("No data found.")
            return

        # Contar os disponíveis e indisponíveis
        available_count = sum(1 for row in data[1:] if ping_ip(row[0]))
        unavailable_count = len(data[1:]) - available_count

        # Gerar HTML com a tabela
        html_table = generate_html_table(data, available_count, unavailable_count)

        # Salvar o resultado em um arquivo HTML
        with open("ip_connectivity_report.html", "w") as file:
            file.write(html_table)

        print("Dashboard de conectividade gerado com sucesso: ip_connectivity_report.html")

    except HttpError as err:
        print(err)

if __name__ == "__main__":
    main()
