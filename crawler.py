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
import win32com.client

excelFilePath = "../vulsList/vulsList.xlsx"

excelapp = win32com.client.Dispatch("Excel.Application")
excelapp.Visible = 0
excelxls = excelapp.Workbooks.Open(excelFilePath)

whiteList = []
blackList = []

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
'''
run.Cells(1, 1).Value = 'Date'
run.Cells(1, 2).Value = 'Title'
run.Cells(1, 3).Value = 'Platform'
'''
data = ['Date', 'Title', 'Platform']
run.Range(run.Cells(line, 1), run.Cells(line, 3)).Value = data
line += 1

begin = datetime(2015, 10, 26)
end = datetime(2015, 10, 28)

urlList = [
	'https://www.exploit-db.com/remote/?order_by=date&order=desc',
	'https://www.exploit-db.com/webapps/?order_by=date&order=desc',
	'https://www.exploit-db.com/local/?order_by=date&order=desc',
	'https://www.exploit-db.com/dos/?order_by=date&order=desc'
]

def getExploitDB(url):
	global line
	s = Session()
	
	req = Request('GET', url)
	prepped = req.prepare()
	prepped.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.99 Safari/537.36' 
	r = s.send(prepped)
	print 'Get : %s' % url
	soup = BeautifulSoup(r.content)
	rows = soup.find("table").find("tbody").findAll("tr")
	date = 0	
	for row in rows:
		cells = row.findAll("td")
		date = datetime.strptime(cells[0].getText(), "%Y-%m-%d")  
		if (begin <= date and date <= end): 
			data = [date, cells[4].getText(), cells[5].getText()]
			run.Range(run.Cells(line, 1), run.Cells(line, 3)).Value = data
			line += 1;
	time.sleep(0.5)
	return date

for url in urlList:
	run.Cells(line, 1).Value = url
	line += 1
	pg = 1;
	while(1):
		pgdate = getExploitDB(url+"&pg="+str(pg))	
		if (pgdate >= begin):
			pg += 1
		else:
			break;



url = 'https://www.hkcert.org/security-bulletin?p_p_id=3tech_list_security_bulletin_full_WAR_3tech_list_security_bulletin_fullportlet&_3tech_list_security_bulletin_full_WAR_3tech_list_security_bulletin_fullportlet_cur='
run.Cells(line, 1).Value = url
line += 1

def getHkcert(url):
	global line
	s = Session()
	req = Request('GET', url)
	prepped = req.prepare()
	prepped.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.99 Safari/537.36' 
	r = s.send(prepped)
	print 'Get : %s' % url

	soup = BeautifulSoup(r.content)
	rows = soup.find("table", attrs={"class": "sdchk_table3"}).find("tbody").findAll("tr")
	date = 0
	for row in rows:
		cells = row.findAll("td")
		date = datetime.strptime(cells[3].getText(), "%Y / %m / %d")  
		if (begin <= date and date <= end): 
			data = [date, cells[1].getText()]
			run.Range(run.Cells(line, 1), run.Cells(line, 2)).Value = data
			line += 1;
	time.sleep(0.5)
	return date

pg = 1;
while(1):
	pgdate = getHkcert(url+str(pg))	
	if (pgdate >= begin):
		pg += 1
	else:
		break;



url = 'http://www.nsfocus.net/index.php?act=sec_bug'
run.Cells(line, 1).Value = url
line += 1

def getNsfocus(url):
	global line
	s = Session()
	req = Request('GET', url)
	prepped = req.prepare()
	prepped.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.99 Safari/537.36' 
	r = s.send(prepped)
	print 'Get : %s' % url

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
			data = [date, title]
			run.Range(run.Cells(line, 1), run.Cells(line, 2)).Value = data
			line += 1;
	time.sleep(0.5)
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
