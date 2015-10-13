# -*- coding: utf-8 -*-
# coding=utf-8

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import string
import time
import csv
from BeautifulSoup import BeautifulSoup
from requests import Request, Session
from datetime import datetime
import opencc

vulsListFile = open('./vulsList.csv', 'wb')
VLWriter = csv.writer(vulsListFile)

data = [['Date', 'Title', 'Platform']]
VLWriter.writerows(data)

begin = datetime(2015, 9, 17)
end = datetime(2015, 9, 24)

urlList = [
	'https://www.exploit-db.com/remote/?order_by=date&order=desc',
	'https://www.exploit-db.com/webapps/?order_by=date&order=desc',
	'https://www.exploit-db.com/local/?order_by=date&order=desc',
	'https://www.exploit-db.com/dos/?order_by=date&order=desc'
]

def getExploitDB(url):
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
			data = [[date, cells[4].getText(), cells[5].getText()]]
			VLWriter.writerows(data)
	time.sleep(0.5)
	return date

for url in urlList:
	VLWriter.writerows([[url]])
	pg = 1;
	while(1):
		pgdate = getExploitDB(url+"&pg="+str(pg))	
		if (pgdate >= begin):
			pg += 1
		else:
			break;


url = 'https://www.hkcert.org/security-bulletin?p_p_id=3tech_list_security_bulletin_full_WAR_3tech_list_security_bulletin_fullportlet&_3tech_list_security_bulletin_full_WAR_3tech_list_security_bulletin_fullportlet_cur='
VLWriter.writerows([[url]])

def getHkcert(url):
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
			data = [[date, cells[1].getText()]]
			VLWriter.writerows(data)
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
VLWriter.writerows([[url]])

def getNsfocus(url):
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
			data = [[date, title]]
			VLWriter.writerows(data)
	time.sleep(0.5)
	return date

pg = 1;
while(1):
	pgdate = getNsfocus(url+"&page="+str(pg))	
	if (pgdate >= begin):
		pg += 1
	else:
		break;

vulsListFile.close()
