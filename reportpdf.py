import requests
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors

print("Welcome to Zabbix Automatic Report Prototype !")
print("==============================================")
hostID = input("Enter Host ID : ")
category = input("Enter Category [CPU | Memory | Disk | System] : ")
print("==============================================")

API_ENDPOINT = "https://slave.msbu.cloud/api_jsonrpc.php"
API_TOKEN = "679f4b4cec4e1fdf02ef7b6024e5ab530fdfe84a4c037bec0b4e3784afa299de"

headers = {
	"Authorization": f"Bearer {API_TOKEN}",
	"Content-Type": "application/json"
	}

payload = {
	"jsonrpc": "2.0",
	"method": "item.get",
	"params": {
		"output": "extend",
		"hostids": f"{hostID}",
		"search": { "name": f"{category}"},
		"sortfield": "name"
	},
	"id": 1
}

response = requests.post(url=API_ENDPOINT, json=payload, headers=headers)
dictResponse = response.json()
# print(dictResponse["result"])

result = []

for i in dictResponse["result"] :
	payload = {
		"jsonrpc": "2.0",
		"method": "trend.get",
		"params": {
			"output": "extend",
			"itemids": f"{i["itemid"]}",
			"limit": "1"
		},
		"id": 1
	}

	response = requests.post(url=API_ENDPOINT, json=payload, headers=headers)
	trendData = response.json()
	
	try :
		a = f"Item ID: {i["itemid"]}; Name: {i["name"]}; Value: {trendData["result"][0]["value_avg"]};"
		result.append(str(a))
	except IndexError:
		a = f"Item ID: {i["itemid"]}; Name: {i["name"]}; Value: -;"
		result.append(str(a))		

fileName = "Report.pdf"
documentTitle = "Report"
title = "Monitoring Report"

pdf = canvas.Canvas(fileName)
pdf.setTitle(documentTitle)

pdf.setFont("Courier-Bold", 20)
pdf.drawCentredString(300, 780, title)

pdf.line(30, 750, 550, 750)

text = pdf.beginText(20, 700)
text.setFont("Courier", 12)
for line in result :
	text.textLine(line)
pdf.drawText(text)

pdf.save()

print("Successfully generated report !")