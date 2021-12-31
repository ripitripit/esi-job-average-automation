import win32com.client as win32
import pandas as pd
import os
import time
import numpy as np
from datetime import datetime
from tkinter import *
from bs4 import BeautifulSoup

#declare outlook variables
outlook=win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

def getEmail(x):
    folder = inbox.Folders[x].Items
    msg = folder.GetLast()
    return(msg)

def scrapeEmails():
    checkoutFolder = input('Checkout Folder Name: ')
    startStr = input('Enter Start Time by Day, Month, Year <01-05-19>: ')
    endStr = input('Enter End Time by Day, Month, Year <31-05-19>: ')
    print('')
    print('Step 1/4 is in progress...\n')
    f = open('ScrapedEmails.html','w+')
    messages = inbox.Folders[checkoutFolder].Items
    startDate = time.strptime(startStr, '%d-%m-%y')
    endDate =  time.strptime(endStr, '%d-%m-%y')
    acceptCt = 0
    omitCt = 0
    for msg in messages:
        msgTime = datetime.strptime(msg.SentOn.strftime("%H:%M:%S"), '%H:%M:%S')
        try:
            if msgTime.hour is 0 and msgTime.minute >= 40 and msgTime.minute <= 60 or msgTime.hour is 1 and msgTime.minute >= 0 and msgTime.minute <= 45: #or msgTime.hour is 2 and msgTime.minute >= 0 and msgTime.minute <= 15:
                msgDate = msg.SentOn.strftime('%d-%m-%y')
                msgDate = time.strptime(msgDate, '%d-%m-%y')
                if msgDate >= startDate and msgDate <= endDate:
                    print(msg.HTMLBody, file=f)
                    print('Accepted result; email details:', msgDate, "; ", msgTime)
                    acceptCt += 1          
            else:
                print('Omitted result')
                omitCt += 1
        except:
            print('Error')
            continue
    f.close()
    print('')
    print('Step 1/4 is complete.')
    print('Total# of emails added to ScrapedEmails.html:', acceptCt)
    print('Total# of emails omitted from ScrapedEmails.html:', omitCt)
    print('')
    
def emails_toCSV():
    print('Step 2/4 is in progress...\n')
    df = pd.read_html('ScrapedEmails.html', skiprows=1)
    count = 0
    appendCt = 0
    omitCt = 0
    list1 = []
    while count < len(df):
        if len(df[count]) > 21:
            list1.append(df[count])
            print('Appended table to list')
            appendCt += 1
        else:
            print('Omitted table')
            omitCt += 1
        count+=1
    df2 = pd.concat(list1)
    print (len(list1))
    df2.to_csv ('ScrapedEmails_toCSV.csv')
    print('Step 2/4 is complete.\n')
    print('Total# of tables appended to list:', appendCt)
    print('Total# of tables omitted from list:', omitCt)
    print('')

def cleanDataframes():
    print('Step 3/4 is in progress...')
    df = pd.read_csv('ScrapedEmails_toCSV.csv', encoding = 'utf8')
    df.drop(['2', '7','9'], axis=1, inplace=True)
    df['Start_Str'] = df['0'] + ' ' + df['3']
    df['End_Str'] = df['0'] + ' ' + df['4']
    df['Start_Time'] = pd.to_datetime(df['Start_Str'], errors='coerce')
    df['End_Time'] = pd.to_datetime(df['End_Str'], errors='coerce')
    df['Day_of_Week'] = df['Start_Time'].dt.day_name()
    df['Run_Time_Abs_Time'] = (df['End_Time'] - df['Start_Time']).abs()
    df['Run_Time_Abs_Seconds'] = (df['End_Time'] - df['Start_Time']).dt.total_seconds().abs()
    df.to_csv('cleanedDataframes.csv')
    print('Step 3/4 is complete.')
    print('Data has been cleaned and exported to cleanedDataframes.csv.\n')

def calcAvgs():
    print('Step 4/4 is in progress...')
    df = pd.read_csv('cleanedDataframes.csv', encoding = 'utf8')
    df = df.groupby(['Day_of_Week', '5'], as_index=False)['Run_Time_Abs_Seconds'].mean()
    df['RunTime_Final'] = (pd.to_timedelta(df['Run_Time_Abs_Seconds'], unit='s')).astype(str)
    df['Total_RunTime_Str'] = df['RunTime_Final'].str[6:15]
    df2 = pd.read_csv('cleanedDataframes.csv', encoding = 'utf8')
    df2['Total_Results_Float'] = 0
    df2['Total_Results_Float'] = df2['8'].astype(float)
    df2 = df2.groupby(['Day_of_Week', '5'], as_index=False)['Total_Results_Float'].mean().round(0)
    df3 = df.merge(df2, how='inner')
    df3.drop(['Run_Time_Abs_Seconds', 'RunTime_Final'], axis=1, inplace=True)
    df3.to_csv('calcAvgs.csv')
    print('Step 4/4 is complete.')
    print('Averages have been calculated and exported to calcAvgs.csv.\n')

def createTemplates():

    #create dataframes from templatess
    sunDf = pd.read_excel('Templates.xlsx', sheet_name = 'Sun', engine='openpyxl')
    monDf = pd.read_excel('Templates.xlsx', sheet_name = 'Tue', engine='openpyxl')
    tueDf = pd.read_excel('Templates.xlsx', sheet_name = 'Wed', engine='openpyxl')
    wedDf = pd.read_excel('Templates.xlsx', sheet_name = 'Thu', engine='openpyxl')
    thuDf = pd.read_excel('Templates.xlsx', sheet_name = 'Fri', engine='openpyxl')
    friDf = pd.read_excel('Templates.xlsx', sheet_name = 'Sat', engine='openpyxl')
    satDf = pd.read_excel('Templates.xlsx', sheet_name = 'Sun', engine='openpyxl')
    
    #create dataframes from the calcAvgs file
    jobsDf = pd.read_csv('calcAvgs.csv', encoding = 'utf8', usecols=['Day_of_Week', '5', 'Total_RunTime_Str', 'Total_Results_Float'])
    avgSunDf = jobsDf[jobsDf['Day_of_Week'].str.contains('Sunday')]
    avgMonDf = jobsDf[jobsDf['Day_of_Week'].str.contains('Monday')]
    avgTueDf = jobsDf[jobsDf['Day_of_Week'].str.contains('Tuesday')]
    avgWedDf = jobsDf[jobsDf['Day_of_Week'].str.contains('Wednesday')]
    avgThuDf = jobsDf[jobsDf['Day_of_Week'].str.contains('Thursday')]
    avgFriDf = jobsDf[jobsDf['Day_of_Week'].str.contains('Friday')]
    avgSatDf = jobsDf[jobsDf['Day_of_Week'].str.contains('Saturday')]
    
    #run a comparison and replace current averages with new metrics
    def replaceAvgs(var1, var2, var3):
        avgTime = ''
        avgResult = ''
        for job2 in var2['5']:
            for job in var1['Critical Process Name']:
                if job.strip() == job2.strip():
                    getRow = var2.loc[(var1['Critical Process Name'] == job.strip()) & var2['Day_of_Week'].str.contains(var3)]
                    avgTime = getRow['Total_RunTime_Str']
                    avgResult = getRow['Total_RunTime_Str']
                    var1.loc['Average'] = avgResult
                    var1.loc['Avg Runtime'] = avgResult         
                    return(var1)
                
    #create html templates w/ new averages      
    with open('Sunday.html', 'w') as file:
        file.write(replaceAvgs(sunDf, avgSunDf, 'Sunday').to_html())
    with open('Monday.html', 'w') as file:
        file.write(replaceAvgs(monDf, avgMonDf, 'Monday').to_html())
    with open('Tuesday.html', 'w') as file:
        file.write(replaceAvgs(tueDf, avgTueDf, 'Tuesday').to_html())
    with open('Wednesday.html', 'w') as file:
        file.write(replaceAvgs(wedDf, avgWedDf, 'Wednesday').to_html())
    with open('Thursday.html', 'w') as file:
        file.write(replaceAvgs(thuDf, avgThuDf, 'Thursday').to_html())
    with open('Friday.html', 'w') as file:
        file.write(replaceAvgs(friDf, avgFriDf, 'Friday').to_html())
    with open('Saturday.html', 'w') as file:
        file.write(replaceAvgs(satDf, avgSatDf, 'Saturday').to_html())


def main():
    #create window object
    window = Tk()
    window.title('ESI Monthly Averages Automation Tool')

    def run(command):
        (str(command))

    #define column 0 buttons
    b1 = Button(window,text="Step 1:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(scrapeEmails()))
    b1.grid(row=1,column=0)
    b1 = Button(window,text="Step 2:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(emails_toCSV()))
    b1.grid(row=2,column=0)
    b1 = Button(window,text="Step 3:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(cleanDataframes()))
    b1.grid(row=3,column=0)
    b1 = Button(window,text="Step 4:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(calcAvgs()))
    b1.grid(row=4,column=0)
    b1 = Button(window,text="Step 4:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(createTemplates()))
    b1.grid(row=5,column=0)

    #define column 1 labels
    L1 = Label(window, text="Steps", width=35)
    L1.grid(row=0,column=1)
    L1 = Label(window, text="Scrape emails from the hourly checkout\nfolder & insert them into an HTML file.", anchor='w', width=35)
    L1.grid(row=1,column=1)
    L1 = Label(window, text="Convert the HTML file to a CSV file.", anchor='w', width=35)
    L1.grid(row=2,column=1)
    L1 = Label(window, text="Read the CSV file\n& clean the data.", anchor='w', width=35)
    L1.grid(row=3,column=1)
    L1 = Label(window, text="Calculate averages\nby day of week.", anchor='w', width=35)
    L1.grid(row=4,column=1)
    L1 = Label(window, text="IDK YET\n.", anchor='w', width=35)
    L1.grid(row=5,column=1)

    #define column 2 labels
    L1 = Label(window, text="File Output", width=20)
    L1.grid(row=0,column=2)
    L1 = Label(window, text="ScrapedEmails.html", anchor='w', width=20)
    L1.grid(row=1,column=2)
    L1 = Label(window, text="ScrapedEmails_toCSV.csv", anchor='w', width=20)
    L1.grid(row=2,column=2)
    L1 = Label(window, text="cleanedDataframes.csv", anchor='w', width=20)
    L1.grid(row=3,column=2)
    L2 = Label(window, text="calcAvgs.csv", anchor='w', width=20)
    L2.grid(row=4,column=2)
    L2 = Label(window, text="IDK_YET.csv", anchor='w', width=20)
    L2.grid(row=5,column=2)

    #declare main loop
    window.mainloop()

main()
