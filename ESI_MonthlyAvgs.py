import win32com.client as win32
#from comtypes.client import CreateObject, GetActiveObject
import pandas as pd
import os
import time
import numpy as np
from tkinter import *

def scrapeEmails():
    checkoutFolder = input('Checkout Folder Name: ')
    startStr = input('Enter Start Time by Day, Month, Year <01-05-19>: ')
    endStr = input('Enter End Time by Day, Month, Year <31-05-19>: ')
    print('')
    print('Step 1/4 is in progress...\n')
    #try:
    outlook=win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    #except:
    #outlook=win32.client.GetActiveObject("Outlook.Application").GetNamespace("MAPI")
    f = open('ScrapedEmails.html','w+')
    inbox=outlook.GetDefaultFolder(6)
    messages=inbox.Folders[checkoutFolder].Items
    startDate = time.strptime(startStr, '%d-%m-%y')
    endDate =  time.strptime(endStr, '%d-%m-%y')
    acceptCt = 0
    omitCt = 0
    for msg in messages:
        try:
            if '12:00 AM' in msg.subject or '12:00 am' in msg.subject or '12:00 AM' in msg.subject or '12:00AM' in msg.subject:
                msgDate = msg.SentOn.strftime('%d-%m-%y')
                msgDate = time.strptime(msgDate, '%d-%m-%y')
                if msgDate >= startDate and msgDate <= endDate:
                    print(msg.HTMLBody, file=f)
                    print('Accepted result; email details:', msgDate)
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
        if len(df[count]) > 25:
            list1.append(df[count])
            print('Appended table to list')
            appendCt += 1
        else:
            print('Omitted table')
            omitCt += 1
        count+=1
    df2 = pd.concat(list1)
    df2.to_csv ('ScrapedEmails_toCSV.csv')
    print('Step 2/4 is complete.\n')
    print('Total# of tables appended to list:', appendCt)
    print('Total# of tables omitted from list:', omitCt)
    print('')


def cleanDataframes():
    print('Step 3/4 is in progress...')
    df = pd.read_csv('ScrapedEmails_toCSV.csv', encoding = 'utf8')
    df.drop(['2', '7','9'], axis=1, inplace=True)
    searchForJobs = ['Rumba', 'Patient Pay', 'Manifest', 'CE2000', 'Medispan', 'Database']
    searchForChars = ['#', '-', '--']
    df = df[~df['5'].str.contains('|'.join(searchForJobs))]
    #df = df[~df['8'].str.contains('|'.join(searchForChars))]
    #df['6'].replace('', np.nan, inplace=True)
    #df = df.dropna(subset=['6'], inplace=True)
    df['Start_Str'] = df['0'] + ' ' + df['3']
    df['End_Str'] = df['0'] + ' ' + df['4']
    df['Start_Time'] = pd.to_datetime(df['Start_Str'])
    df['End_Time'] = pd.to_datetime(df['End_Str'])
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
    df2 = df2.groupby(['Day_of_Week', '5'], as_index=False)['Total_Results_Float'].mean()
    df3 = df.merge(df2, how='inner')
    df3.drop(['Run_Time_Abs_Seconds', 'RunTime_Final'], axis=1, inplace=True)
    df3.to_csv('calcAvgs.csv')
    print('Step 4/4 is complete.')
    print('Averages have been calculated and exported to calcAvgs.csv.\n')

def main():
    #create window object
    window = Tk()
    window.title('ESI Monthly Averages Automation Tool')

    def run(command):
        (str(command))

    #define buttons
    b1 = Button(window,text="Step 1:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(scrapeEmails()))
    b1.grid(row=1,column=0)
    b1 = Button(window,text="Step 2:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(emails_toCSV()))
    b1.grid(row=2,column=0)
    b1 = Button(window,text="Step 3:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(cleanDataframes()))
    b1.grid(row=3,column=0)
    b1 = Button(window,text="Step 4:", width=12, height=2, bg='cyan', fg='black', command=lambda: run(calcAvgs()))
    b1.grid(row=4,column=0)

    #define labels
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

    window.mainloop()

main()
