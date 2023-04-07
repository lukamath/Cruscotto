# Importing pandas
import os
import sys
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotview import view
import saveashtml
import schedule
import time

def salvataggiintegrati():
	#salvataggi integrati
	saveashtml.salvacode()
	saveashtml.salvaprod()
	url = "data_source/code.html"
	table = pd.read_html(url)[0]
	table.to_excel("data_source/code.xlsx")
	
#scrittura corretta della fascia
now = datetime.now()    
today=now.strftime("%d/%m/%Y")
later=now+timedelta(hours=1)
current_time = now.strftime("%H:%M:%S")
later_time=later.strftime("%H:%M:%S")
timeframe=now.strftime("%H")+ ":00-" + str(later.strftime("%H")) +":00"
hour=now.strftime("%H")
inthour=int(hour)

plt.ion()
fig = plt.figure()
ax1 = fig.add_subplot(111)

forecast = pd.read_excel('data_source/forecast.xlsx', sheet_name="x")
dayforecast = forecast[['fcst_name', 'Gestione CLT','timeframe',today]].copy()
dayforecast["hour"]=dayforecast["timeframe"].str[6:8].astype(int)
dayforecast.to_excel("service/dayforecast.xlsx")
# currentforecast e' il mio forecast fino all'ora corrente
currentforecast=dayforecast.loc[dayforecast["hour"]<=inthour]   

listacode = pd.read_excel('data_source/forecast.xlsx', sheet_name="Legenda Inbound")
themap=pd.read_excel('data_source/map.xlsx', sheet_name="Map")
minimap=themap.groupby(['fcst_name', 'Report Activity'], as_index=False).sum()
fcstmap=minimap[['fcst_name', 'Report Activity']].copy()

currentforecast=currentforecast.merge(fcstmap, on='fcst_name')
currentforecast.to_excel("service/currentforecast.xlsx")

salvataggiintegrati()
# url = "data_source/code.html"
# table = pd.read_html(url)[0]
# table.to_excel("data_source/code.xlsx") 

#   *** dataframe for charts ***    #
#dayforecastonfcst=dayforecast.loc[dayforecast['fcst_name']!='COV-NOTTE']
dayforecastonfcst=dayforecast.merge(fcstmap, on='fcst_name')
dayforecastonfcst.to_excel('service/dayforecastonfcst.xlsx')
dayforecastonactivity=dayforecastonfcst.groupby(['Report Activity', 'hour'], as_index=False).sum()
dayforecastonactivity.to_excel('service/dayforecastonactivity.xlsx')

dayforecastmob=dayforecastonfcst.loc[dayforecastonfcst["Report Activity"]=='MOB']
dayforecastmob=dayforecastmob.loc[dayforecastonfcst["Gestione CLT"]=='POST']
dayforecastmob=dayforecastmob.groupby('hour').sum() 
dayforecastmob.to_excel('output/dayforecastmob.xlsx')
#   ----------------------------    #

#-----------------------------------------------------------------------------------
#code = pd.read_excel('data_source/code.xlsx',skiprows=2)  #per salvataggio di Rende
#code.rename(columns={"&nbsp": "vag"}, inplace=True)

#code = pd.read_excel('data_source/code.xlsx',skiprows=[1, 3])
code = pd.read_excel('data_source/code.xlsx',skiprows=1)
code.rename(columns={"Unnamed: 0_level_1": "vag"}, inplace=True)
#-----------------------------------------------------------------------------------

match=themap.merge(code, left_on='Coda',right_on='vag')
match.rename(columns={"Riclassifica": "fcst_name"}, inplace=True)
match.to_excel("service/match&timeframe.xlsx")

currentforecastonfcst=currentforecast.groupby('fcst_name').sum()
currentforecastonactivity=currentforecast.groupby('Report Activity').sum()
currentforecastonfcst["hour"]=inthour		#fill the column 'hour' with the current hour
currentforecastonactivity["hour"]=inthour	#fill the column 'hour' with the current hour
currentforecastonfcst.to_excel("service/currentforecastonfcst.xlsx")
currentforecastonactivity.to_excel("service/currentforecastonactivity.xlsx")

matchonfcst=match.groupby('fcst_name').sum()
matchonactivity=match.groupby('Report Activity').sum()
matchonfcst.to_excel("service/matchonfcst.xlsx")
matchonactivity.to_excel("service/matchonactivity.xlsx")

reporthouronfcst=matchonfcst.merge(currentforecastonfcst,left_on='fcst_name',right_on='fcst_name')
reporthouronfcst['Delta_Offerto']=(reporthouronfcst['Offerte']/reporthouronfcst[today] - 1)*100
reporthouronfcst=reporthouronfcst.drop(['Unnamed: 0', 'T.A.','Livello di Servizio %','% Cleared','hour'], axis=1) #export rende
reporthouronfcst.to_excel('output/reporthouronfcst.xlsx')
reporthouronfcst = reporthouronfcst.cumsum()

reporthouronactivity=matchonactivity.merge(currentforecastonactivity,left_on='Report Activity', right_on='Report Activity')
reporthouronactivity['Delta_Offerto']=(reporthouronactivity['Offerte']/reporthouronactivity[today] - 1)*100
reporthouronactivity=reporthouronactivity.drop(['Unnamed: 0', 'T.A.','Livello di Servizio %','% Cleared','hour'], axis=1) #export rende
reporthouronactivity.to_excel('output/reporthouronactivity.xlsx')

#copy of current situation on activity
#pasthouronactivity=reporthouronactivity.copy()
#pasthouronactivity.to_excel('service/pasthouronactivity.xlsx')
pasthouronactivity = pd.read_excel('service/pasthouronactivity.xlsx')

#fig = plt.figure()

def realtime_data():
	salvataggiintegrati()
	code = pd.read_excel('data_source/code.xlsx',skiprows=1)
	code.rename(columns={"Unnamed: 0_level_1": "vag"}, inplace=True)
	match=themap.merge(code, left_on='Coda',right_on='vag')
	match.rename(columns={"Riclassifica": "fcst_name"}, inplace=True)
	match.to_excel("service/match&timeframe.xlsx")

	matchonfcst=match.groupby('fcst_name').sum()
	matchonactivity=match.groupby('Report Activity').sum()
	matchonfcst.to_excel("service/matchonfcst.xlsx")
	matchonactivity.to_excel("service/matchonactivity.xlsx")

	reporthouronfcst=matchonfcst.merge(currentforecastonfcst,left_on='fcst_name',right_on='fcst_name')
	reporthouronfcst['Delta_Offerto']=(reporthouronfcst['Offerte']/reporthouronfcst[today] - 1)*100
	reporthouronfcst=reporthouronfcst.drop(['Unnamed: 0', 'T.A.','Livello di Servizio %','% Cleared','hour'], axis=1) #export rende
	reporthouronfcst.to_excel('output/reporthouronfcst.xlsx')
	reporthouronfcst = reporthouronfcst.cumsum()

	reporthouronactivity=matchonactivity.merge(currentforecastonactivity,left_on='Report Activity', right_on='Report Activity')
	reporthouronactivity['Delta_Offerto']=(reporthouronactivity['Offerte']/reporthouronactivity[today] - 1)*100
	reporthouronactivity=reporthouronactivity.drop(['Unnamed: 0', 'T.A.','Livello di Servizio %','% Cleared','hour'], axis=1) #export rende
	reporthouronactivity.to_excel('output/reporthouronactivity.xlsx')

	pasthouronactivity = pd.read_excel('service/pasthouronactivity.xlsx')

	deltahouronactivity=reporthouronactivity.subtract(pasthouronactivity)
	deltahouronactivity.to_excel('service/deltahouronactivity.xlsx')

def saveprevious():
	#copy of current situation on activity
	pasthouronactivity=reporthouronactivity.copy()
	pasthouronactivity.to_excel('service/pasthouronactivity.xlsx')

def save_framereport():
	now = datetime.now()    
	minframe=now.strftime("%H-%M")
	currentdate=now.strftime("%Y-%m-%d")
	framereport=reporthouronactivity.copy()
	framereport.to_excel('output/framereport/code_'+currentdate+"-"+minframe+ '.xlsx')

def refresh_chart():
	now = datetime.now()  
	plt.clf()
	fig.set_figwidth(13)
	ax1 = fig.add_subplot(1,1,1) #questa riga mi serve anche per cancellare eventuale linea riplottata su stesso timeframe 
	#ax1.set_xlim(8, 24)
	dayrealmob=pd.read_excel('service/dayrealmob.xlsx', usecols=[1,2])
	dayrealmob.rename(columns={today: "Offerte"}, inplace=True)
	reporthouronactivity=pd.read_excel('output/reporthouronactivity.xlsx')
    #indexmob = reporthouronactivity[reporthouronactivity['Report Activity'] == 'MOB'].index
	indexmob = 3
	indexmobhour = dayrealmob[dayrealmob['hour'] == int(now.strftime("%H"))+1].index 
	#indexmobhour=10
	i=reporthouronactivity['Offerte'][indexmob]-pasthouronactivity['Offerte'][indexmob]
	j=reporthouronactivity['Offerte'][indexmob]
	thismin=now.strftime("%M")
	print('this minute is: ' + thismin)
	print('inthour: ' + str(inthour))
	print('indexmobhour: ' + str([indexmobhour]))
	print('previous: ' + str(pasthouronactivity['Offerte'][indexmob]))
	print('current: ' + str(reporthouronactivity['Offerte'][indexmob]))
	print(j)
	print(i)
	#print(reporthouronactivity['Offerte'][indexmob+1])
	#dayrealmob=dayrealmob.replace(i,j)
	dayrealmob['Offerte'][indexmobhour]=i
	#dayrealmob.loc['Offerte',indexmobhour]=j   #vedi.loc 
	dayrealmob.to_excel('service/dayrealmob.xlsx')

	dayrealmob=dayrealmob.set_index('hour')
	ax1.plot(dayforecastmob[today])
	ax1.plot(dayrealmob['Offerte'])
	plt.title('Mobile POST')
	ax1.legend(['forecast','real'])
	fig.canvas.draw()
	fig.canvas.flush_events()
	plt.pause(0.013)
	
#vedi --> https://schedule.readthedocs.io/en/stable/examples.html
schedule.every(30).seconds.do(refresh_chart)
schedule.every(13).seconds.do(realtime_data)
schedule.every().hour.at(":00").do(saveprevious) # Run saveprevious every hour at the 00 minute
schedule.every().hour.at(":00").do(save_framereport) 
schedule.every().hour.at(":10").do(save_framereport) 
schedule.every().hour.at(":20").do(save_framereport) 
schedule.every().hour.at(":30").do(save_framereport) 
schedule.every().hour.at(":40").do(save_framereport) # Run saveprevious every hour at the 00 minute
schedule.every().hour.at(":50").do(save_framereport) # Run saveprevious every hour at the 00 minute


while True:
    schedule.run_pending()
    time.sleep(2)