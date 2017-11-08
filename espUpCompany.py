#!/usr/bin/python
# coding=utf-8  
import re
import time
from bs4 import BeautifulSoup
import requests
import pandas
import numpy as np
from openpyxl import load_workbook


def func_getHistoryValue(stockNum):
	dfs = pandas.read_html('https://finance.google.com/finance/historical?q=TPE:' + stockNum + '&start=0&num=200')
	history = dfs[2]
	history = history.ix[:,4]
	avg5a = []
	avg20a = []
	avg5 = 0.0
	avg20 = 0.0
	for index in range(1,len(history)):
		avg5 = avg5 + float(history[index])
		avg20 = avg20 + float(history[index])
		if(index > 19):
			avg20v = avg20/20
			avg20a.append(avg20v)
			avg20 = avg20 - float(history[index-19])
		if(index > 4):
			avg5v = avg5/5
			avg5a.append(avg5v)
			avg5 = avg5 - float(history[index-4])
	max = avg5a[0]
	min = avg5a[0]
	cou = 0
	cod = 0
	buy = u'等'
	for index in range(1,len(history)):
		if(max > avg5a[index]):
			max = avg5a[index]
			cou = cou + 1
		else:
			break;

	for index in range(1,len(history)):
		if(min < avg5a[index]):
			min = avg5a[index]
			cod = cod + 1
		else:
			break;

	if(avg5a[0] > avg20a[0]):
		print ('stock Up!!!')
		if(abs(avg5a[0] - avg20a[0]) < (float(history[1])*0.02)):
			status = u'上(5接近20)'+str(cou)+":"+str(cod)
			if(abs(avg5a[0] - float(history[1])) < (float(history[1])*0.02)):
				if(cou >= 1):
					buy = u'買'
				else:
					buy = u'賣'
			else:
				buy = u'等'
#		elif(abs(avg5a[0] - float(history[1])) < (float(history[1])*0.02)):
#			status = u'上(5接近現價)'+str(cou)+":"+str(cod)
#			if(cou >= 3):
#				buy = u'買'
		else:
			if(cod >= 3):
				buy = u'賣'
			status = u'上'+str(cou)+":"+str(cod)
	else:
		print ('stock Down!!!')
		print ('sell')
		if(abs(avg5a[0] - avg20a[0]) < (float(history[1])*0.02)):
			status = u'下(5接近20)'+str(cou)+":"+str(cod)
			if(cod >= 3):
				buy = u'賣'
#		elif(abs(avg5a[0] - float(history[1])) < (float(history[1])*0.02)):
#			status = u'下(5接近現價)'+str(cou)+":"+str(cod)
#			if(cod >= 3):
#				buy = u'賣'
		else:
			if(cou >= 3):
				buy = u'買'
			status = u'下'+str(cou)+":"+str(cod)
	return (status,buy)

def func_ConnectToKimoGiven(stockNum):
	isOK = 0
	err_cou = 0
	url = 'https://tw.stock.yahoo.com/d/s/dividend_' + stockNum + '.html'
	#print (url)
	while isOK == 0:
		try:
			dfs = pandas.read_html(url)
			skg = dfs[9]
			#print (skg[1][1],skg[2][1])
			isOK = 1
		except Exception as e: 
			print(e)
			err_cou = err_cou + 1 
			print ('except func_ConnectToKimoGiven')
			if (err_cou > 5):
				print ('err too much')
				isOK = 1
	return (skg[1][1],skg[2][1])

def search_company( stockNum , factory, val):
	try:
		dfs = pandas.read_html('https://tw.stock.yahoo.com/d/s/company_' + stockNum[0:4] + '.html')
		scd = dfs[9]
		scd = scd.replace(np.nan, '', regex=True)
		scd = scd.replace(u'元', '', regex=True)
		sum1 = 0.0
		sum2 = 0.0
		for index in range(1,5):
			sum1 = sum1 + float(scd[3][index])
			sum2 = sum2 + float(scd[5][index])
		if (sum1 > float(scd[5][1])) and (sum1 > (sum2/4)) and (sum1 > 0) and (float(scd[5][1]) > 0) and (sum2 > 0):
			(status,buy) = func_getHistoryValue(stockNum[0:4])
			(money, stock) = func_ConnectToKimoGiven(stockNum[0:4])	
			print (u'股號:',stockNum[0:4],u'股名:',stockNum[4:],u'  產業別:',factory,u'  價位:',val,u'  前四季:',sum1,u'  去年:',float(scd[5][1]),u'  四年均:',sum2/4,u'成長率:',str((sum1-float(scd[5][1]))/sum1*100),'%',u'  配息:', money , u'  配股:' , stock , u'現金殖利率' ,float(money)/float(val)*100,'%',u'線型',status,u'購買狀態',buy)
			return (stockNum[0:4] + ',' +stockNum[4:] + ',' + factory + ',' + val + ','+ str(sum1) + ',' + str(float(scd[5][1])) + ',' + str(sum2/4) +','+ str((sum1-float(scd[5][1]))/sum1*100)+'%' +',' + money + ',' + stock +',' + str(float(money)/float(val)*100) + '%,'+status+ ','+buy+' \n')
		else:	
			return ''
	except Exception as e: 
		print(e)
		return ''


def func_search_stockvalue(stockNum):
	try:
		res = requests.get('https://tw.stock.yahoo.com/q/q?s=' + stockNum[0:4])
		bs = BeautifulSoup(res.text,'lxml')
		tables = bs.select('table')
		trs = tables[6].select('tr')
		tds = trs[1].select('td')
		#print (tds[0].text[0:len(tds[0].text)-6],tds[2].text)
		value = tds[2].text
	except:
		print ("請確定網路或股號是否正確!!")
		value = ''
	return value


res = requests.get('http://isin.twse.com.tw/isin/C_public.jsp?strMode=2')
#print (res.text)
bs = BeautifulSoup(res.text,'lxml')
# table[1]/tbody/tr/td/table[2]
#index = 0
tables = bs.select('table')
#for table in tables:
#	print (index)
#	print (table)
#	index = index + 1

trs = tables[1].select('tr')
index = 0
arrs = []
narrs =[]
now = time.strftime("%Y%m%d%H%M")
filename = 'epsUp-' + now + '.csv'
f = open(filename, 'w')
try:
	f.write(u'股號' + u',股名' +u',產業別' + u',價位' + u',前四季EPS' + u',去年EPS' + u',四年均EPS' + u',成長率'+ u',配息'+ u',配股' + u',現金殖利率' + u',線型(2%)'+',購買狀態\n')
	for tr in trs:
		#print (index)
		tds = tr.select('td')
		if(len(tds) > 2):
			# 将正则表达式编译成Pattern对象
			pattern = re.compile(r'[1-9](\d+)\s')
			# 使用Pattern匹配文本，获得匹配结果，无法匹配时将返回None
			match = pattern.match(tds[0].text)
			if match:
				# 使用Match获得分组信息
				m = match.group()
				if(len(m) == 5):
					#print (match.group())
					print (tds[0].text)
					arrs.append(tds[0].text)
					narrs.append(tds[4].text)
		index = index + 1

	for index in range(0,len(arrs)):
		#print (arr)
		value = func_search_stockvalue(arrs[index])
		factory = narrs[index]
		if(value != ''):
			f.write(search_company(arrs[index],factory,value))
except Exception as e: 
	print(e)

f.close()