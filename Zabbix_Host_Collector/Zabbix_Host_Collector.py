from openpyxl import Workbook, load_workbook
from zabbix_utils import ZabbixAPI
import pandas as pd

cliente="cliente"


def coletar_host(cliente):

    #Conecta na API do Zabbix
    api = ZabbixAPI(url="zabbixurl")
    api.login(user="user", password="password")

    customer= cliente

    # Coleta os nomes dos hosts pela API
    hosts = api.host.get(output=['name'])


    filtered_hosts = [host['name'] for host in hosts if customer in host['name']]

    df_hosts = pd.DataFrame(filtered_hosts, columns=['Host'])

    df_hosts['Host'] = df_hosts['Host'].str.replace(fr'\({customer}\)', '', regex=True).str.strip()
    df_hosts = df_hosts.sort_values(by='Host')


    df_hosts.to_excel("Hosts_Coletados.xlsx", index=False)
    print("Arquivo 'Hosts_Coletados.xlsx' criado com sucesso.")

coletar_host(cliente)