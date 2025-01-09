import requests,json,csv,codecs,datetime,time
import xlsxwriter
from spire.xls import *
from spire.xls.common import *
from metrics import metrics

ApiUrl = '<api_url>'
token = "<token>"
header = {"Content-Type":"application/json", "Authorization": f"Bearer {token}"}

alphabet = list(map(chr, range(ord('A'), ord('Z')+1)))

# x=(datetime.datetime.now()-datetime.timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
# y=(datetime.datetime.now()).strftime("%Y-%m-%d %H:%M:%S")

x = datetime.date(2024, 12, 1).strftime("%Y-%m-%d %H:%M:%S")
y = datetime.date(2025, 12, 31).strftime("%Y-%m-%d %H:%M:%S")

def formatTimestamp(x,y):
  p=time.strptime(x,"%Y-%m-%d %H:%M:%S")
  starttime = str(int(time.mktime(p)))
  q=time.strptime(y,"%Y-%m-%d %H:%M:%S")
  endtime = str(int(time.mktime(q)))

  return starttime,endtime

timestamp = formatTimestamp(x, y)

def get_groupid(groupname):
    data = {
        "jsonrpc": "2.0",
        "method": "hostgroup.get",
        "params": {
            "output": "extend"
        },
        "id": 1
    }
    getgroupid=requests.post(url=ApiUrl,headers=header,json=data)
    groupids=json.loads(getgroupid.content)["result"]
    groupid = [groupids[i]['groupid'] for i in range(len(groupids)) if groupname in groupids[i]['name']][0]
    return groupid

def get_hosts(groupname):
  groupid = get_groupid(groupname)
  data ={
    "jsonrpc": "2.0",
      "method": "host.get",
      "params": {
      "output": [
              "hostid",
              "host",
              "groups"
          ],
    "groupids": groupid,
    "selectHosts": [],
    "selectInterfaces": [
                "ip",
                "interfaceid"
            ]
    },
    "id": 1
  }
  gethost=requests.post(url=ApiUrl,headers=header,json=data)

  return json.loads(gethost.content)["result"]

def get_Templates(groupName,hostid):
  data = {
    "jsonrpc": "2.0",
      "method": "template.get",
      "params": {
      "output": [
              "hostid",
              "host"
          ],
      "search":{
          "groupname": groupName
      },
    "selectHosts": []
    },
    "id": 2
  }
  gettemplate=requests.post(url=ApiUrl,headers=header,json=data)
  result = json.loads(gettemplate.content)["result"]
  
  templates = [result[i]['host'] for i in range(len(result)) if {"hostid":hostid} in result[i]['hosts']]

  return templates

def get_item(hostid, itemkey):
    data={
            "jsonrpc": "2.0",
            "method": "item.get",
            "params": {
                "output": [
                    "itemid",
                    "name"
                ],
                "searchWildcardsEnabled": "true",
                "search": {
                    "key_": itemkey
                },
                "hostids": hostid
            },
            "id": 3
        }
    getitem=requests.post(url=ApiUrl,headers=header,json=data)
    item=json.loads(getitem.content)['result']
    
    return item

def get_trend_data(itemid,timestamp):
    trenddata={
        "jsonrpc":"2.0",
        "method":"trend.get",
        "params":{
            "output": [
                "itemid",
                "value_max",
                "value_avg",
                "value_min"
                "clock"
            ],
            "time_from":timestamp[0],
            "time_till":timestamp[1],
            "itemids": itemid,
            "limit":1
        },
        
        "id":4
        }
    gettrend=requests.post(url=ApiUrl,headers=header,json=trenddata)
    trend=json.loads(gettrend.content)['result']
    return trend

def createreport(groupName):
        fname = f'{groupName}_Report_monthly.xlsx'
        workbook = xlsxwriter.Workbook(fname)
        # hostgroup format
        format_hostgroup = workbook.add_format()
        format_hostgroup.set_bold
        format_hostgroup.set_font_size(14)

        # Title format
        format_title = workbook.add_format()
        format_title.set_border(1)  # border
        format_title.set_bg_color('#808080')  # background color
        format_title.set_align('center')
        format_title.set_bold()
        format_title.set_valign('vcenter')
        format_title.set_font_size(12)

        # Date format
        format_date = workbook.add_format()
        format_date.set_border(1)
        format_date.set_align('center')
        format_date.set_valign('vcenter')
        format_date.set_font_size(12)
        format_date.set_num_format('yyyy-mm-dd')
        format_date.set_bg_color('#D3D3D3')

        # history data format
        format_value = workbook.add_format()
        format_value.set_border(1)
        format_value.set_align('center')
        format_value.set_valign('vcenter')
        format_value.set_font_size(12)

        format_percentage = workbook.add_format({'num_format': '0.00'})
        format_percentage.set_border(1)
        format_percentage.set_align('center')
        format_percentage.set_valign('vcenter')
        
        format_saturation = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#90EE90",
        })
        format_traffic = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#FFFFED",
        })
        format_latency = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#ADD8E6",
        })
        format_errors = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#FF7F7F",
        })

        format_srel = {
            "Saturation": format_saturation,
            "Traffic": format_traffic,
            "Latency": format_latency,
            "Errors": format_errors,
        }


        # get all host in group
        hosts = get_hosts(groupName)
        # print(hosts)
        # hosts = [{'hostid': '10681', 'host': 'artemis', 'interfaces': [{'interfaceid': '69', 'ip': '192.168.122.14'}, {'interfaceid': '79', 'ip': '192.168.122.14'}]}]

        for host in hosts:
            hostnames = host['host']
            print(hostnames)
            hostid = host['hostid']
            templates = get_Templates(groupName,hostid)
            sheet = workbook.add_worksheet(hostnames)
            sheet.set_column('A:A', 15)
            sheet.set_column('B:Z', 25)
            dayInMonth = 31
            i = 1

            # Write per Category
            for category in metrics["Linux by Zabbix agent"]:
                metricCount = len(metrics["Linux by Zabbix agent"][category])
                sheet.merge_range(f"B{i}:{alphabet[metricCount]}{i}", category, format_srel[category])

                for t in range(1, dayInMonth+1):
                    sheet.write(f'A{i+1}', "Date", format_title)
                    sheet.write(f'A{i+t+1}', f'2024-12-{t}', format_date)


                if "Linux by Zabbix agent" not in templates:
                    break
                
                j = 1
                for metric in metrics["Linux by Zabbix agent"][category]:

                    item = get_item(hostid, metric["key"])
                    itemname = metric["name"]

                    try:
                        itemid = item[0]['itemid']
                    except IndexError:
                        itemid = "-"

                    sheet.write(f'{alphabet[j]}{i+1}', f'{itemname}', format_title)

                    # set colum width
                    trendData = []
                    if itemid == "-":
                        j+=1
                        continue
                    
                    for t in range(1, dayInMonth+1):
                        
                        timestamp = formatTimestamp(datetime.datetime(2024, 12, t).strftime("%Y-%m-%d %H:%M:%S"), datetime.datetime(2025, 12, t, 23, 59, 59).strftime("%Y-%m-%d %H:%M:%S"))
                        data = get_trend_data(itemid,timestamp)
                        # print(data, timestamp)
                        
                        try:
                            trendData.append(data[0]['value_avg'])  
                        except IndexError:
                            trendData.append(0.0)

                    # print(itemname)
                    # print("trends",trendData)
                    k=2
                    for d in trendData:
                        sheet.write_number(f'{alphabet[j]}{i+k}', float(d), format_percentage)
                        k+=1
                    j+=1
                i+=dayInMonth+3


                # chart = workbook.add_chart({'type': 'line'})
                # chart.set_title({'name': 'CPU Utilization Monthly'})
                # chart.add_series({
                #     "values": f"='{hostnames}'!$B$2:$B$32", 
                #     "name": f"{itemname}",
                #     "line": {"color": "#1ac6c0"}
                #     })
                # sheet.insert_chart('C1', chart)


            # write title


            

        # # create Linux workbook
        # worksheet1 = workbook.add_worksheet("Linux Servers")
        # # set colum width
        # worksheet1.set_column('A:A', 51)
        # worksheet1.set_column('B:G', 15)
        # worksheet1.set_column('H:I', 20)
        # # set row high
        # worksheet1.set_default_row(25)
        # # Freeze title
        # worksheet1.freeze_panes(4, 1)

        # # write hostgroup
        # worksheet1.write(0, 0, f"{hostsgroup}", format_hostgroup)
        # # write title
        # i = 0
        # for title in Title1:
        #     worksheet1.write(3, i, title, format_title)
        #     i += 1
        # # write host history data
        # j = 4
        # hosts=get_hosts(get_linux_hosts(token),token)
        # # print(hosts)
        # for host in hosts:
        #     hostid=host['hostid']
        #     hostname=host['name']
        #     try:
        #         hostip=host['interfaces'][0]['ip']
        #     except IndexError:
        #         hostip=""
        #     row=get_linux_host_hist(hostid,hostname,hostip,token,timestamp)
        #     # print("/n row /n",row)


        #     worksheet1.write(j, 0, row[0]['Hostname'], format_value)
        #     worksheet1.write(j, 1, row[0]['IP'], format_value)
        #     try:
        #         worksheet1.write_number(j, 2, float(row[0]['CPU Utilization avg5']), format_percentage)
        #     except ValueError:
        #         worksheet1.write(j, 2, '-', format_value)
        #     try:
        #         worksheet1.write_number(j, 3, float(row[0]['Memory Utilization']), format_percentage)
        #     except ValueError:
        #         worksheet1.write(j, 3, '-', format_value)
        #     try:
        #         worksheet1.write_number(j, 4, float(row[0]['Space Utilization']), format_percentage)
        #     except ValueError:
        #         worksheet1.write(j, 4, '-', format_value)
        #     try:
        #         worksheet1.write_number(j, 5, int(row[0]['Traffic In']), format_value)
        #     except ValueError:
        #         worksheet1.write(j, 5, '-', format_value)
        #     try:
        #         worksheet1.write_number(j, 6, int(row[0]['Traffic Out']), format_value)
        #     except ValueError:
        #         worksheet1.write(j, 6, '-', format_value)
        #     worksheet1.write(j, 7, row[0]['Start Time'], format_value)
        #     worksheet1.write(j, 8, row[0]['End Time'], format_value)
        #     j += 1
        
        # # create Windows workbook
        # worksheet2 = workbook.add_worksheet("Windows Servers")
        # # set colum width
        # worksheet2.set_column('A:A', 51)
        # worksheet2.set_column('B:G', 15)
        # worksheet2.set_column('H:I', 20)
        # # set row high
        # worksheet2.set_default_row(25)
        # # Freeze title
        # worksheet2.freeze_panes(1, 1)
        # # write title
        # i = 0
        # for title in Title2:
        #     worksheet2.write(0, i, title, format_title)
        #     i += 1
        # write host history data
#         j = 1
#         hosts=get_hosts(get_windows_hosts(token),token)
#         for host in hosts:
#             hostid=host['hostid']
#             hostname=host['name']
#             hostip=host['interfaces'][0]['ip']
#             row=get_windows_host_hist(hostid,hostname,hostip,token,timestamp)
# #            print(row)

#             worksheet2.write(j, 0, row[0]['Hostname'], format_value)
#             worksheet2.write(j, 1, row[0]['IP'], format_value)
#             worksheet2.write_number(j, 2, float(row[0]['CPU Utilization']), format_percentage)
#             worksheet2.write_number(j, 3, float(row[0]['Memory Utilization']), format_percentage)
#             worksheet2.write_number(j, 4, float(row[0]['Space Utilization']), format_percentage)
#             worksheet2.write_number(j, 5, int(row[0]['Traffic In']), format_value)
#             worksheet2.write_number(j, 6, int(row[0]['Traffic Out']), format_value)
#             worksheet2.write(j, 7, row[0]['Start Time'], format_value)
#             worksheet2.write(j, 8, row[0]['End Time'], format_value)
#             j += 1
        
        
        workbook.close()


createreport("PT. ABC")