from zabbix_utils import ZabbixAPI
import csv

api = ZabbixAPI(url="zabbixurl")
api.login(user="user", password="password")

print(api.api_version())

cliente = "cliente" #colocar aqui o nome do cliente
#na planilha host_teste.csv colocar o host e o ip dos host que deseja cadastrar
def create_host(host,ip):
    try:
        create_host= api.host.create({
            "host": host,
            "name": f"({cliente}) {host}",
            "groups": [{"groupid":31}], #colocar id do hostgroup do cliente
            "interfaces": [{
                "type": 1,
                "main": 1,
                "useip": 1,
                "ip": ip,
                "dns": "",
                "port": "10050",
            }],
            "templates":[
                {"templateid": "10564"} #templateid ICMP Ping
    #            {"templateid": "10081"} #templateid Windows by Zabbix agent
            ],
            "monitored_by": 1,
            "proxy_hostid": "10846", #sintaxe versao 6.4 #colocar o proxyid conforme o do cliente (para descobrir id: proxy.get)
            "status": 1 #status desabilitado
        })
        print(f'Host cadastrado com sucesso! {host}')
    except Exception as err:
        print(f'falha ao cadastrar o host: erro {err}')


with open("hosts_teste.csv","r") as arquivo:
    arquivo_csv = csv.reader(arquivo,delimiter=';')
    for linha in arquivo_csv:
        hostname = linha[0]
        ipaddress = linha[1]
        create_host(hostname,ipaddress)
