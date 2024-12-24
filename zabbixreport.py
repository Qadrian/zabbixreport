import requests,json,csv,codecs,datetime,time
import xlsxwriter

ApiUrl = 'https://zabbix.msbu.cloud/api_jsonrpc.php'
user="Admin"
password="Jakarta2023@"
token="679f4b4cec4e1fdf02ef7b6024e5ab530fdfe84a4c037bec0b4e3784afa299de"
header = {"Content-Type":"application/json", "Authorization": f"Bearer {token}"}

Title1=['Hostname','IP','CPU Utilization','Memory Util (%)','Space Util (/ %)','Net In (Bits)','Net Out (Bits)','Start Time','End Time']
Title2=['Hostname','IP','CPU Util (%)','Memory Util (%)','Space Util (C: %)','Net In (Bits)','Net Out (Bits)','Start Time','End Time']
x=(datetime.datetime.now()-datetime.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
y=(datetime.datetime.now()).strftime("%Y-%m-%d %H:%M:%S")

def gettoken():
    data = {"jsonrpc": "2.0",
                "method": "user.login",
                "params": {
                    "username": user,
                    "password": password
                },
                "id": 1,
                "auth": None
            }
    auth=requests.post(url=ApiUrl,headers=header,json=data)
    print(auth)
    return token

def timestamp(x,y):
    p=time.strptime(x,"%Y-%m-%d %H:%M:%S")
    starttime = str(int(time.mktime(p)))
    q=time.strptime(y,"%Y-%m-%d %H:%M:%S")
    endtime = str(int(time.mktime(q)))

    print("start: ", x, " end: ", y)
    return starttime,endtime

def logout(auth):
    data={
            "jsonrpc": "2.0",
            "method": "user.logout",
            "params": [],
            "id": 1
            }
    auth=requests.post(url=ApiUrl,headers=header,json=data)
    return json.loads(auth.content)

def get_hosts_by_group(hostids, auth, groupid="35"):
    data = {
        "jsonrpc": "2.0",
        "method": "host.get",
        "params": {
            "output": ["hostid", "name"],
            "hostids": hostids,
            "groupids": groupid,
            "search":{
                "status": "0"
import requests,json,csv,codecs,datetime,time
import xlsxwriter

ApiUrl = 'https://zabbix.msbu.cloud/api_jsonrpc.php'
user="Admin"
password="Jakarta2023@"
token="679f4b4cec4e1fdf02ef7b6024e5ab530fdfe84a4c037bec0b4e3784afa299de"
header = {"Content-Type":"application/json", "Authorization": f"Bearer {token}"}

Title1=['Hostname','IP','CPU Utilization','Memory Util (%)','Space Util (/ %)','Net In (Bits)','Net Out (Bits)','Start Time','End Time']
Title2=['Hostname','IP','CPU Util (%)','Memory Util (%)','Space Util (C: %)','Net In (Bits)','Net Out (Bits)','Start Time','End Time']
x=(datetime.datetime.now()-datetime.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
y=(datetime.datetime.now()).strftime("%Y-%m-%d %H:%M:%S")

def gettoken():
    data = {"jsonrpc": "2.0",
                "method": "user.login",
                "params": {
                    "username": user,
                    "password": password
                },
                "id": 1,
                "auth": None
            }
    auth=requests.post(url=ApiUrl,headers=header,json=data)
    print(auth)
    return token

def timestamp(x,y):
    p=time.strptime(x,"%Y-%m-%d %H:%M:%S")
    starttime = str(int(time.mktime(p)))
    q=time.strptime(y,"%Y-%m-%d %H:%M:%S")
    endtime = str(int(time.mktime(q)))

    print("start: ", x, " end: ", y)
    return starttime,endtime

def logout(auth):
    data={
            "jsonrpc": "2.0",
            "method": "user.logout",
            "params": [],
            "id": 1
            }
    auth=requests.post(url=ApiUrl,headers=header,json=data)
    return json.loads(auth.content)

def get_hosts_by_group(hostids, auth, groupid="35"):
    data = {
        "jsonrpc": "2.0",
        "method": "host.get",
        "params": {
            "output": ["hostid", "name"],
            "hostids": hostids,
            "groupids": groupid,
            "search":{
                "status": "0"
            },
            "selectInterfaces": [
                "ip",
                "interfaceid"
            ],
        },
        "id": 1
    }
    gethost = requests.post(url=ApiUrl, headers=header, json=data)
    return json.loads(gethost.content)["result"]

def get_linux_hosts(auth):
    data ={
              "jsonrpc": "2.0",
              "method": "template.get",
              "params": {
                  "output": [
                      "hostid",
                      "host"
                  ],
                  "search": {
                      "template": [
                          "Template OS Linux by Zabbix agent",
                      ]
                  },
                  "selectHosts": [],
              },
            "id": 1
        }
    gethost=requests.post(url=ApiUrl,headers=header,json=data)
    hostids=[]
    linux_hosts=json.loads(gethost.content)["result"]
    for hostlist in linux_hosts:
        for hostid in hostlist['hosts']:
            hostids.append(hostid['hostid'])
    return hostids

def get_windows_hosts(auth):
    data ={

