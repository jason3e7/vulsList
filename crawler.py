# -*- coding: utf8 -*-
# coding=utf-8

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import string
import time
import csv
from BeautifulSoup import BeautifulSoup
from requests import Request, Session

vulsListFile = open('./vulsList.csv', 'wb')
VLWriter = csv.writer(vulsListFile)

data = [['Date', 'Title', 'Platform']]
VLWriter.writerows(data)

urlList = [
	'https://www.exploit-db.com/remote/',
	'https://www.exploit-db.com/webapps/',
	'https://www.exploit-db.com/local/',
	'https://www.exploit-db.com/dos/'
	]

begin = "2015-09-17"
end = "2015-09-24"

for url in urlList:
	VLWriter.writerows([[url]])
	s = Session()
	req = Request('GET', url)
	prepped = req.prepare()
	prepped.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2454.99 Safari/537.36' 
	r = s.send(prepped)
	print 'Get : %s' % url

	soup = BeautifulSoup(r.content)
	rows = soup.find("table").find("tbody").findAll("tr")
	
	for row in rows:
		cells = row.findAll("td")
		date = cells[0].getText()
		if (begin <= date and date <= end): 
			data = [[date, cells[4].getText(), cells[5].getText()]]
			VLWriter.writerows(data)
			
	time.sleep(0.5)

vulsListFile.close()
