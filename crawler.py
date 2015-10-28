# -*- coding: utf-8 -*-
# coding=utf-8

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import string
import time
from BeautifulSoup import BeautifulSoup
from requests import Request, Session
from datetime import datetime
import opencc
import re
import win32com.client

sTime = 0.1

excelFilePath = "../vulsList/vulsList.xlsx"

excelapp = win32com.client.Dispatch("Excel.Application")
excelapp.Visible = 0
excelxls = excelapp.Workbooks.Open(excelFilePath)

AllCVEList = []
whiteList = []
blackList = []

print "Read xlsx file"

wl = excelxls.Worksheets("whiteList")
used = wl.UsedRange
nrows = used.Row + used.Rows.Count

for i in range(2, nrows):
	whiteList.append(str(wl.Cells(i, 1)))

bl = excelxls.Worksheets("blackList")
used = bl.UsedRange
nrows = used.Row + used.Rows.Count

for i in range(2, nrows):
	blackList.append(str(bl.Cells(i, 1)))

#print whiteList
#print blackList

line = 1
run = excelxls.Worksheets('run')
data = ['Date', 'Title', 'Platform', 'Source', 'CVE', 'Status']
run.Range(run.Cells(line, 1), run.Cells(line, 6)).Value = data
line += 1

begin = datetime(2015, 10, 26)
end = datetime(2015, 10, 28)

def getHttp(url):
	time.sleep(sTime)
	s = Session()	
	req = Request('GET', url)
	prepped = req.prepare()
	prepped.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.99 Safari/537.36' 
	print 'Get : %s' % url
	return s.send(prepped)

def checkCVE(cve):
	global AllCVEList
	if re.match('CVE-\d{4}-\d{4,7}', cve) is None:
		return "NoCVE"
	if cve in AllCVEList:
		return "CVErepeat"
	AllCVEList.append(cve)
	return "New"

def inBlackList(title):
	global blackList
	for black in blackList:
		if re.search(black, title, re.IGNORECASE):
			return 1
	return 0

urlList = [
	'https://www.exploit-db.com/remote/?order_by=date&order=desc',
	'https://www.exploit-db.com/webapps/?order_by=date&order=desc',
	'https://www.exploit-db.com/local/?order_by=date&order=desc',
	'https://www.exploit-db.com/dos/?order_by=date&order=desc'
]

def getExploitDB(url):
	global line
	r = getHttp(url)
	soup = BeautifulSoup(r.content)
	rows = soup.find("table").find("tbody").findAll("tr")
	date = 0	
	for row in rows:
		cells = row.findAll("td")
		date = datetime.strptime(cells[0].getText(), "%Y-%m-%d")  
		if (begin <= date and date <= end):
			title = cells[4].getText()
			status = ""
			if inBlackList(title):
				status = status + "black,"
			source = cells[4].find('a').get('href')	
			sourceR = getHttp(source)
			sourceHttp = BeautifulSoup(sourceR.content)
			tdList = sourceHttp.find("table", {"class" : "exploit_list"}).findAll("td")
			cve = tdList[1].getText()
			cve = cve.replace(":", "-")
			data = [date, title, cells[5].getText(), source, cve, status + checkCVE(cve)]
			run.Range(run.Cells(line, 1), run.Cells(line, 6)).Value = data
			line += 1;
	return date

for url in urlList:
	pg = 1;
	while(1):
		pgdate = getExploitDB(url+"&pg="+str(pg))	
		if (pgdate >= begin):
			pg += 1
		else:
			break;

url = 'https://www.hkcert.org/security-bulletin?p_p_id=3tech_list_security_bulletin_full_WAR_3tech_list_security_bulletin_fullportlet&_3tech_list_security_bulletin_full_WAR_3tech_list_security_bulletin_fullportlet_cur='

def getHkcert(url):
	global line
	r = getHttp(url)
	soup = BeautifulSoup(r.content)
	rows = soup.find("table", attrs={"class": "sdchk_table3"}).find("tbody").findAll("tr")
	date = 0
	for row in rows:
		cells = row.findAll("td")
		date = datetime.strptime(cells[3].getText(), "%Y / %m / %d")  
		if (begin <= date and date <= end):
			title = cells[1].getText()
			statusData = ""
			if inBlackList(title):
				statusData = statusData + "black,"
			source = 'https://www.hkcert.org/' + str(cells[1].find('a').get('href'))
			sourceR = getHttp(source)
			sourceHttp = BeautifulSoup(sourceR.content)
			cveList = sourceHttp.find("div", {"id" : "content6"}).findAll("li")
			cveData = ""
			statusData = ""
			for cve in cveList:
				cveNum = cve.getText()
				cveData = cveData + cveNum + ","
				statusData = statusData + checkCVE(cveNum) + ","
			data = [date, cells[1].getText(), "", source, cveData, statusData]
			run.Range(run.Cells(line, 1), run.Cells(line, 6)).Value = data
			line += 1;
	return date

pg = 1;
while(1):
	pgdate = getHkcert(url+str(pg))	
	if (pgdate >= begin):
		pg += 1
	else:
		break;


url = 'http://www.nsfocus.net/index.php?act=sec_bug'

def getNsfocus(url):
	global line
	r = getHttp(url)
	r.encoding = r.apparent_encoding
	soup = BeautifulSoup(r.text)
	rows = soup.find("ul", attrs={"class": "vul_list"}).findAll("li")
	cc = opencc.OpenCC('s2t')
	date = 0
	for row in rows:
		# cn word print ERROR but save file OK
		date = datetime.strptime(row.find("span").getText(), "%Y-%m-%d")  
		if (begin <= date and date <= end):
			# save utf8 tw use excel import OK
			title = cc.convert(row.find("a").getText())
			status = ""
			if inBlackList(title):
				status = status + "black,"
			m = re.search('CVE-\d{4}-\d{4,7}', title)
			cve = m.group(0)
			source = "http://www.nsfocus.net" + str(row.find("a").get("href"))
			data = [date, title[:-15], "", source, cve, status + checkCVE(cve)]
			run.Range(run.Cells(line, 1), run.Cells(line, 6)).Value = data
			line += 1;
	return date

pg = 1;
while(1):
	pgdate = getNsfocus(url+"&page="+str(pg))	
	if (pgdate >= begin):
		pg += 1
	else:
		break;


excelxls.Save()
excelapp.Quit()
