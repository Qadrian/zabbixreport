import requests,json,csv,codecs,datetime,time
import xlsxwriter

ApiUrl = ''
user=""
password="
token=""
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
              "jsonrpc": "2.0",
              "method": "template.get",
              "params": {
                  "output": [
                      "hostid",
                      "host"
                  ],
                  "search": {
                      "host": [
                          "Template OS Windows by Zabbix agent active",
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

def get_linux_host_hist(hostid,hostname,hostip,auth,timestamp):
    host=[]
    item1=[]
    item2=[]
    dic1={}
    for j in ['system.cpu.util','vm.memory.size[pavailable]','vfs.fs.dependent.size[/,pused]','net.if.in[\"e*\"]','net.if.out[\"e*\"]']:
        data={
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "output": [
                    "itemid"
                ],
                "searchWildcardsEnabled": "true",
                "search": {
                    "key_": j
                },
                "hostids": hostid
            },
            "id": 1
        }
        getitem=requests.post(url=ApiUrl,headers=header,json=data)
        item=json.loads(getitem.content)['result']

        try:
            itemid2=item[0]['itemid']
        except IndexError:
            break

        hisdata={
            "jsonrpc":"2.0",
            "method":"history.get",
            "params":{
                "output":"extend",
                "time_from":timestamp[0],
                "time_till":timestamp[1],
                "history":0,
                "sortfield": "clock",
                "sortorder": "DESC",
                "itemids": '%s' %(item[0]['itemid']),
                "limit":1
            },
            "id":1
            }
        get_host_hist=requests.post(url=ApiUrl,headers=header,json=hisdata)
        hist=json.loads(get_host_hist.content)['result']
        item1.append(hist)

    for j in ['system.cpu.util','vm.memory.size[pavailable]','vfs.fs.dependent.size[/,pused]','net.if.in[\"e*\"]','net.if.out[\"e*\"]']:
        data={
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "output": [
                    "itemid"
                ],
                "searchWildcardsEnabled": "true",
                "search": {
                    "key_": j
                },
                "hostids": hostid
            },
            "id": 1
        }
        getitem=requests.post(url=ApiUrl,headers=header,json=data)
        item=json.loads(getitem.content)['result']

        try:
            itemid2=item[0]['itemid']
        except IndexError:
            break

        trendata={
            "jsonrpc":"2.0",
            "method":"trend.get",
            "params":{
                "output": [
                    "itemid",
                    "value_max",
                    "value_avg"
                ],
                "time_from":timestamp[0],
                "time_till":timestamp[1],
                "itemids": '%s' %(item[0]['itemid']),
                "limit":1
            },
            "id":1
            }
        gettrend=requests.post(url=ApiUrl,headers=header,json=trendata)
        trend=json.loads(gettrend.content)['result']
        item2.append(trend)

    dic1['Hostname']=hostname
    dic1['IP']=hostip
    try:
        dic1['CPU Utilization avg5']=item2[0][0]['value_avg']
    except IndexError:
        dic1['CPU Utilization avg5']=0
    try:
        dic1['Memory Utilization']=item2[1][0]['value_avg']
    except IndexError:
        dic1['Memory Utilization']=0
    try:
        dic1['Space Utilization']=item2[2][0]['value_avg']
    except IndexError:
        dic1['Space Utilization']=0
    try:
        dic1['Traffic In']=item2[3][0]['value_avg']
    except IndexError:
        dic1['Traffic In']=0
    try:
        dic1['Traffic Out']=item2[4][0]['value_avg']
    except IndexError:
        dic1['Traffic Out']=0
    dic1['Start Time']=x
    dic1['End Time']=y
    host.append(dic1)
    if not host:
        print("no record")
    else:
        return host

def get_windows_host_hist(hostid,hostname,hostip,auth,timestamp):
    host=[]
    item1=[]
    item2=[]
    dic1={}
    for j in ['system.cpu.util','vm.memory.util','vfs.fs.size[C:,pused]','net.if.in["Amazon Elastic Network Adapter"]','net.if.out["Amazon Elastic Network Adapter"]']:
        data={
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "output": [
                    "itemid"
                ],
                "searchWildcardsEnabled": "true",
                "search": {
                    "key_": j
                },
                "hostids": hostid
            },
            "id": 1
        }
        getitem=requests.post(url=ApiUrl,headers=header,json=data)
        item=json.loads(getitem.content)['result']

        try:
            itemid2=item[0]['itemid']
        except IndexError:
            break

        hisdata={
            "jsonrpc":"2.0",
            "method":"history.get",
            "params":{
                "output":"extend",
                "time_from":timestamp[0],
                "time_till":timestamp[1],
                "history":0,
                "sortfield": "clock",
                "sortorder": "DESC",
                "itemids": '%s' %(item[0]['itemid']),
                "limit":1
            },
            "id":1
            }
        get_host_hist=requests.post(url=ApiUrl,headers=header,json=hisdata)
        hist=json.loads(get_host_hist.content)['result']
        item1.append(hist)

    for j in ['system.cpu.util','vm.memory.util','vfs.fs.size[C:,pused]','net.if.in["Amazon Elastic Network Adapter"]','net.if.out["Amazon Elastic Network Adapter"]']:
        data={
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "output": [
                    "itemid"
                ],
                "searchWildcardsEnabled": "true",
                "search": {
                    "key_": j
                },
                "hostids": hostid
            },
            "id": 1
        }
        getitem=requests.post(url=ApiUrl,headers=header,json=data)
        item=json.loads(getitem.content)['result']

        try:
            itemid2=item[0]['itemid']
        except IndexError:
            break

        trendata={
            "jsonrpc":"2.0",
            "method":"trend.get",
            "params":{
                "output": [
                    "itemid",
                    "value_max",
                    "value_avg"
                ],
                "time_from":timestamp[0],
                "time_till":timestamp[1],
                "itemids": '%s' %(item[0]['itemid']),
                "limit":1
            },
            "id":1
            }
        gettrend=requests.post(url=ApiUrl,headers=header,json=trendata)
        trend=json.loads(gettrend.content)['result']
        item2.append(trend)

    dic1['Hostname']=hostname
    dic1['IP']=hostip
    try:
        dic1['CPU Utilization']=item2[0][0]['value_avg']
    except IndexError:
        dic1['CPU Utilization']=0
    try:
        dic1['Memory Utilization']=item2[1][0]['value_avg']
    except IndexError:
        dic1['Memory Utilization']=0
    try:
        dic1['Space Utilization']=item2[2][0]['value_avg']
    except IndexError:
        dic1['Space Utilization']=0
    try:
        dic1['Traffic In']=item2[3][0]['value_avg']
    except IndexError:
        dic1['Traffic In']=0
    try:
        dic1['Traffic Out']=item2[4][0]['value_avg']
    except IndexError:
        dic1['Traffic Out']=0
    dic1['Start Time']=x
    dic1['End Time']=y
    host.append(dic1)
    if not host:
        print("no record")
    else:
        return host

def createreport():
    fname = 'Zabbix_Report_weekly.xlsx'
    workbook = xlsxwriter.Workbook(fname)
    
    format_title = workbook.add_format()
    format_title.set_border(1)
    format_title.set_bg_color('#1ac6c0')
    format_title.set_align('center')
    format_title.set_bold()
    format_title.set_valign('vcenter')
    format_title.set_font_size(12)
    
    format_value = workbook.add_format()
    format_value.set_border(1)
    format_value.set_align('center')
    format_value.set_valign('vcenter')
    format_value.set_font_size(12)

    format_percentage = workbook.add_format()
    format_percentage.set_border(1)
    format_percentage.set_align('center')
    format_percentage.set_valign('vcenter')
    format_percentage.set_font_size(12)
    format_percentage.set_num_format('0.00')

    # Linux worksheet
    worksheet1 = workbook.add_worksheet("Linux Servers")
    worksheet1.set_column('A:A', 51)
    worksheet1.set_column('B:G', 15)
    worksheet1.set_column('H:I', 20)
    worksheet1.set_default_row(25)
    worksheet1.freeze_panes(1, 1)
    
    i = 0
    for title in Title1:
        worksheet1.write(0, i, title, format_title)
        i += 1
    
    j = 1
    linux_hosts = get_hosts_by_group(get_linux_hosts(token), token)
    for host in linux_hosts:
        hostid = host['hostid']
        hostname = host['name']
        try:
            hostip = host['interfaces'][0]['ip']
        except IndexError:
            hostip = ""
        row = get_linux_host_hist(hostid, hostname, hostip, token, timestamp)
        if row:
            worksheet1.write(j, 0, row[0]['Hostname'], format_value)
            worksheet1.write(j, 1, row[0]['IP'], format_value)
            worksheet1.write_number(j, 2, float(row[0]['CPU Utilization avg5']), format_percentage)
            worksheet1.write_number(j, 3, float(row[0]['Memory Utilization']), format_percentage)
            worksheet1.write_number(j, 4, float(row[0]['Space Utilization']), format_percentage)
            worksheet1.write_number(j, 5, int(row[0]['Traffic In']), format_value)
            worksheet1.write_number(j, 6, int(row[0]['Traffic Out']), format_value)
            worksheet1.write(j, 7, row[0]['Start Time'], format_value)
            worksheet1.write(j, 8, row[0]['End Time'], format_value)
            j += 1

    # Windows worksheet
    worksheet2 = workbook.add_worksheet("Windows Servers")
    worksheet2.set_column('A:A', 51)
    worksheet2.set_column('B:G', 15)
    worksheet2.set_column('H:I', 20)
    worksheet2.set_default_row(25)
    worksheet2.freeze_panes(1, 1)
    
    i = 0
    for title in Title2:
        worksheet2.write(0, i, title, format_title)
        i += 1
    
    j = 1
    windows_hosts = get_hosts_by_group(get_windows_hosts(token), token)
    for host in windows_hosts:
        hostid = host['hostid']
        hostname = host['name']
        try:
            hostip = host['interfaces'][0]['ip']
        except IndexError:
            hostip = ""
        row = get_windows_host_hist(hostid, hostname, hostip, token, timestamp)
        if row:
            worksheet2.write(j, 0, row[0]['Hostname'], format_value)
            worksheet2.write(j, 1, row[0]['IP'], format_value)
            worksheet2.write_number(j, 2, float(row[0]['CPU Utilization']), format_percentage)
            worksheet2.write_number(j, 3, float(row[0]['Memory Utilization']), format_percentage)
            worksheet2.write_number(j, 4, float(row[0]['Space Utilization']), format_percentage)
            worksheet2.write_number(j, 5, int(row[0]['Traffic In']), format_value)
            worksheet2.write_number(j, 6, int(row[0]['Traffic Out']), format_value)
            worksheet2.write(j, 7, row[0]['Start Time'], format_value)
            worksheet2.write(j, 8, row[0]['End Time'], format_value)
            j += 1

    workbook.close()

# Main execution
token = gettoken()
timestamp = timestamp(x, y)
print(timestamp)
createreport()
logout(token)