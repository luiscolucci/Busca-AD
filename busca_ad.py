import pandas as pd
from ldap3 import Server, Connection, ALL, NTLM
import csv
import sys

# --- CONFIGURAÇÃO - PREENCHA COM AS INFORMAÇÕES DO SEU AMBIENTE ---

# Endereço do seu servidor Active Directory (ex: 'ad.suaempresa.com' ou o endereço IP)
AD_SERVER = 'seu_servidor_ad.seu_dominio.com'

# Usuário e senha para conectar ao AD.
# Este usuário precisa de, no mínimo, permissão de leitura no Active Directory.
# IMPORTANTE: Para maior segurança, evite deixar a senha diretamente no código em ambientes
# de produção. Considere usar variáveis de ambiente ou um cofre de segredos.
AD_USER = 'DOMINIO\\seu_usuario'  # Ex: 'SUAEMPRESA\\leitor.ad'
AD_PASSWORD = 'sua_senha_aqui'

# O caminho completo (Distinguished Name - DN) para a OU onde os computadores estão.
# Altere 'seu_dominio' e 'com' para corresponder exatamente ao seu domínio.
SEARCH_BASE_DN = 'OU=Computadores,OU=Equipamentos,DC=seu_dominio,DC=com'

# Nome do seu arquivo CSV de entrada que contém os nomes dos computadores
INPUT_CSV_FILE = 'lista_endpoints.csv'

# Nome exato da coluna no seu CSV que contém os nomes dos computadores
ENDPOINT_COLUMN = 'endpoint'

# Nome do arquivo de saída que será gerado com o resultado
OUTPUT_REPORT_FILE = 'relatorio_verificacao_ad.csv'

# --- FIM DA CONFIGURAÇÃO ---


def main():
    """
    Função principal que orquestra a leitura do CSV, a conexão com o AD
    e a geração do relatório de verificação.
    """
    print("Iniciando o processo de verificação de computadores no Active Directory...")

    # 1. Tenta ler a lista de computadores do arquivo CSV
    try:
        print(f"Lendo o arquivo de entrada: '{INPUT_CSV_FILE}'...")
        df = pd.read_csv(INPUT_CSV_FILE)
        if ENDPOINT_COLUMN not in df.columns:
            print(f"ERRO CRÍTICO: A coluna '{ENDPOINT_COLUMN}' não foi encontrada no arquivo CSV.")
            sys.exit(1)
        
        # Remove espaços em branco, ignora valores nulos e cria a lista
        computer_names = df[ENDPOINT_COLUMN].dropna().str.strip().tolist()
        print(f"Encontrados {len(computer_names)} nomes de computadores para verificar.")
    except FileNotFoundError:
        print(f"ERRO CRÍTICO: O arquivo de entrada '{INPUT_CSV_FILE}' não foi encontrado.")
        sys.exit(1)
    except Exception as e:
        print(f"ERRO CRÍTICO ao ler o arquivo CSV: {e}")
        sys.exit(1)

    # 2. Estabelece a conexão com o Active Directory
    server = Server(AD_SERVER, get_info=ALL)
    conn = None
    try:
        print(f"Conectando ao servidor AD '{AD_SERVER}'...")
        # Usando autenticação NTLM, comum em ambientes Windows
        conn = Connection(server, user=AD_USER, password=AD_PASSWORD, authentication=NTLM, auto_bind=True)
        print("Conexão com o AD estabelecida com sucesso.")
    except Exception as e:
        print(f"ERRO CRÍTICO: Falha ao conectar ou autenticar no Active Directory: {e}")
        sys.exit(1)

    # 3. Itera sobre cada computador, verifica sua existência e armazena os resultados
    results = []
    print(f"\nIniciando a busca na OU: '{SEARCH_BASE_DN}'...")

    for computer_name in computer_names:
        if not computer_name:
            continue  # Pula nomes de computador vazios

        # O filtro de busca procura por um objeto do tipo 'computer' com o nome correspondente.
        # O '$' no final do sAMAccountName é o padrão para contas de computador.
        search_filter = f'(&(objectClass=computer)(sAMAccountName={computer_name}$))'

        # Executa a busca no AD
        conn.search(
            search_base=SEARCH_BASE_DN,
            search_filter=search_filter,
            attributes=['sAMAccountName']  # Só precisamos saber se existe
        )

        # Verifica se a busca retornou alguma entrada (se encontrou o computador)
        if conn.entries:
            status = "OK"
            print(f"  - {computer_name}: {status}")
        else:
            status = "Não Encontrado"
            print(f"  - {computer_name}: {status}")

        results.append({'Computador': computer_name, 'Status': status})

    # Fecha a conexão com o AD
    conn.unbind()
    print("\nConexão com o Active Directory foi encerrada.")

    # 4. Escreve o relatório final em um novo arquivo CSV
    try:
        print(f"Gerando relatório final em '{OUTPUT_REPORT_FILE}'...")
        with open(OUTPUT_REPORT_FILE, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Computador', 'Status']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            writer.writeheader()
            writer.writerows(results)
        print("\nProcesso concluído! Relatório gerado com sucesso!")
    except Exception as e:
        print(f"ERRO CRÍTICO ao gerar o arquivo de relatório: {e}")

if __name__ == '__main__':
    main()