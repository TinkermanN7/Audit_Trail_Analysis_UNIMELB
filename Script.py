#!/usr/bin/env python
# coding: utf-8

# In[1]:


#Import packages

import pandas as pd
import datetime
import xlsxwriter
import math

from datetime import date


# In[2]:


# Necessary week generator function

def weekGenerator(date):
    this_day = date.weekday()
    monday = date - datetime.timedelta(days = this_day)
    dates = [(monday + datetime.timedelta(days=i)).strftime('%d/%m/%Y') for i in range(5)]
    return dates


# In[3]:


# Read excel files.

file1 = "AuditTrailProd.xlsx"
database_xlsx_dict = pd.read_excel(file1, sheet_name = None)


# In[4]:


# Dictionary Keys

database_dict_keys = list(database_xlsx_dict.keys())


# In[5]:


# Read into dataframe

database_df = database_xlsx_dict[database_dict_keys[0]]


# In[35]:


# (Optional)
# How database looks like.

# pd.set_option("display.max_rows", 50,"display.max_columns", None)
# database_df


# In[7]:


# Total changes per username, total changes per type, total changes per object
# Overall Analysis Code

username_changes = {}
type_changes = {}
object_changes = {}

for row in database_df.iterrows():

    Username = row[1]['UserName']
    Type = row[1]['Type']
    Object = row[1]['Object']
    
    if Username not in username_changes:
        username_changes[Username] = 1
    else:
        username_changes[Username] += 1
    
    if Type not in type_changes:
        type_changes[Type] = 1
    else:
        type_changes[Type]+=1
    
    if Object not in object_changes:
        object_changes[Object] = 1
    else:
        object_changes[Object] += 1


# In[8]:


# Write Overall Changes to excel file.

workbook = xlsxwriter.Workbook('Overall_Changes.xlsx')
sheet1 = workbook.add_worksheet()

sheet1.write("A1", "UserName")
sheet1.write("B1", "Changes Made")

count = 1

for user in username_changes:
    sheet1.write(count, 0, user)
    sheet1.write(count, 1, username_changes[user])
    count += 1

sheet2 = workbook.add_worksheet()
sheet2.write("A1", "Type Of Change")
sheet2.write("B1", "Changes Made")

count = 1

for type in type_changes:
    sheet2.write(count, 0, type)
    sheet2.write(count, 1, type_changes[type])
    count += 1
    
sheet3 = workbook.add_worksheet()

sheet3.write("A1", "Object of Change")
sheet3.write("B1", "Changes Made")


count = 1

for object in object_changes:
    sheet3.write(count, 0, object)
    sheet3.write(count, 1, object_changes[object])
    count += 1

    
username_changes = sorted(username_changes.items(), key = lambda x: x[1], reverse = True)
sheet1.write("D2", username_changes[0][0])


# In[9]:


# Daily Analysis Code

today = date.today()
today = str(today)

curr_date = today[-2] + today[-1] + '/' + today[-5] + today[-4] + '/' + today[-10] + today[-9] + today[-8] + today[-7]

# Set current date and previous date to artificial date for testing purposes.
prev_date = (datetime.datetime.today() - datetime.timedelta(days = 1)).strftime('%d/%m/%Y')
curr_date = (datetime.datetime.today()).strftime('%d/%m/%Y')

username_changes = {}
type_changes = {}
object_changes = {}

for row in database_df.iterrows():
    if curr_date not in row[1]['DateTime']:
        continue
    Username = row[1]['UserName']
    Type = row[1]['Type']
    Object = row[1]['Object']
    
    if Username not in username_changes:
        username_changes[Username] = 1
    else:
        username_changes[Username] += 1
    
    if Type not in type_changes:
        type_changes[Type] = 1
    else:
        type_changes[Type]+=1
    
    if Object not in object_changes:
        object_changes[Object] = 1
    else:
        object_changes[Object] += 1


# Write Current Day Changes to excel file.

sheet4 = workbook.add_worksheet()

sheet4.write("A1", "UserName")
sheet4.write("B1", "Changes Made today")

count = 1

for user in username_changes:
    sheet4.write(count, 0, user)
    sheet4.write(count, 1, username_changes[user])
    count += 1

sheet5 = workbook.add_worksheet()
sheet5.write("A1", "Type Of Change")
sheet5.write("B1", "Changes Made today")

count = 1

for type in type_changes:
    sheet5.write(count, 0, type)
    sheet5.write(count, 1, type_changes[type])
    count += 1
    
sheet6 = workbook.add_worksheet()

sheet6.write("A1", "Object of Change")
sheet6.write("B1", "Changes Made today")


count = 1

for object in object_changes:
    sheet6.write(count, 0, object)
    sheet6.write(count, 1, object_changes[object])
    count += 1


# In[10]:


# Weekly Analysis Code

# Hard Code Week for Demonstration Purposes
week = weekGenerator(datetime.datetime.today())


username_changes = {}

for row in database_df.iterrows():

    Username = row[1]['UserName']
    Type = row[1]['Type']
    Object = row[1]['Object']
    date = row[1]['DateTime']
    
    if date[:10] not in week:
        continue
    
    if Username not in username_changes:
        username_changes[Username] = 1
    else:
        username_changes[Username] += 1
    

    
    
deletion = {}
addition = {}
update = {}

for row in database_df.iterrows():
    date = row[1]['DateTime']
    Type = row[1]['Type']
    
    if date[:10] not in week:
        continue
    
    if Type == 'Delete':
        if date[:10] not in deletion:
            deletion[date[:10]] = 1
        else:
            deletion[date[:10]] += 1
            
    if Type == 'Add':
        if date[:10] not in addition:
            addition[date[:10]] = 1
        else:
            addition[date[:10]] += 1
    if Type == 'Update':
        if date[:10] not in update:
            update[date[:10]] = 1
        else:
            update[date[:10]] += 1
    

        
# Write Type Changes for week to file.

sheet7 = workbook.add_worksheet()

sheet7.write("A1", "Date")
sheet7.write("B1", "Deletions Made")

count = 1

for date in deletion:
    sheet7.write(count, 0, date)
    sheet7.write(count, 1, deletion[date])
    count += 1

sheet8 = workbook.add_worksheet()
sheet8.write("A1", "Date")
sheet8.write("B1", "Additions Made")

count = 1

for date in addition:
    sheet8.write(count, 0, date )
    sheet8.write(count, 1, addition[date])
    count += 1
    
sheet9 = workbook.add_worksheet()

sheet9.write("A1", "Date")
sheet9.write("B1", "Updates Made")


count = 1

for date in update:
    sheet9.write(count, 0, date)
    sheet9.write(count, 1, update[date])
    count += 1


    
    
sheet10 = workbook.add_worksheet()

sheet10.write("A1", "UserName")
sheet10.write("B1", "Changes Made")


count = 1

for user in username_changes:
    sheet10.write(count, 0, user)
    sheet10.write(count, 1, username_changes[user])
    count += 1
    
    
workbook.close()
    


# In[11]:


# Find out smallest earliest month and earliest year.

min_month = 13
min_year = 3000

for row in database_df.iterrows():
    year = int(row[1]['DateTime'][6:10])
    if year < min_year:
        min_year = year

# print(min_year)

for row in database_df.iterrows():
    month = int(row[1]['DateTime'][3:5])
    year = int(row[1]['DateTime'][6:10])
    
    if year == min_year and month < min_month:
        min_month = month
        
# print(min_month)

        


# In[12]:


# Find earliest date

earliest_day = 32
for row in database_df.iterrows():
    day = int(row[1]['DateTime'][:2])
    month = int(row[1]['DateTime'][3:5])
    year = int(row[1]['DateTime'][6:10])
    
    if month == min_month and year == min_year and day < earliest_day:
        earliest_day = day
        
first_date = ''
if len(str(earliest_day)) != 2:
    first_date += '0'
    first_date += str(earliest_day)
else:
    first_date+=str(earliest_day)

first_date += '/'


if len(str(min_month)) != 2:
    first_date += '0'
    first_date+= str(min_month)
else:
    first_date+=str(min_month)
    

first_date += '/'
first_date += str(min_year)

# print(first_date)


# In[15]:


database_df


# In[16]:


# Compartmentalise data on a weekly basis

d1 = datetime.date(int(first_date[-4:]), int(first_date[3:5]), int(first_date[:2]))

weeks = {}

for row in database_df.iterrows():
    date_str = row[1]['DateTime'][:10]
    day = int(date_str[:2])
    month = int(date_str[3:5])
    year = int(date_str[6:10])
    
    d2 = datetime.date(year, month, day)
    week = int(math.floor(((d2-d1).days)/7)+1)
    
    Username = row[1]['UserName']
    Type = row[1]['Type']
    Object = row[1]['Object']
    
    if week not in weeks:
        weeks[week] = {'Total_Changes': 1, 'Username_Changes':{Username: 1}, 'Type_Changes': {Type : 1}, 'Object_Changes': {Object : 1}}
        weeks[week]['Week Starting'] = ''
    
    else:
        if Username not in weeks[week]['Username_Changes']:
            weeks[week]['Username_Changes'][Username] = 1
        else:
            weeks[week]['Username_Changes'][Username] += 1
        
        if Type not in weeks[week]['Type_Changes']:
            weeks[week]['Type_Changes'][Type] = 1
        else:
            weeks[week]['Type_Changes'][Type] += 1

        if Object not in weeks[week]['Object_Changes']:
            weeks[week]['Object_Changes'][Object] = 1
        else:
            weeks[week]['Object_Changes'][Object] += 1
            
        weeks[week]['Total_Changes'] += 1
        

# Sort Weeks
weeks = dict(sorted(weeks.items(), key = lambda x: x[0]))


# In[18]:


# Index weeks by starting date of the week:
# YEARS = ['2019', '2020', '2021']

d1 = datetime.date(int(first_date[-4:]), int(first_date[3:5]), int(first_date[:2]))

for week in weeks:
    if str(week) != 'Week Starting':
        week_start_date = d1 + datetime.timedelta(days = 7 * (week-1))
        weeks[week]['Week Starting'] = week_start_date
        
        


# In[26]:


# Compartmentalise data in a montly basis

MONTHS = {1 : 'January', 2 : 'February', 3: 'March', 4 : 'April', 5: 'May', 6: 'June', 7 : 'July', 8 : 'August', 9: 'September', 10: 'October', 11 : 'November', 12 : 'December'}

months = {}
for row in database_df.iterrows():
    date_str = row[1]['DateTime'][:10]
    month = MONTHS[int(date_str[3:5])] + ' ' + date_str[6:10]
    
    Username = row[1]['UserName']
    Type = row[1]['Type']
    Object = row[1]['Object']
    
    
    if month not in months:
        months[month] = {'Total_Changes': 1, 'Username_Changes':{Username: 1}, 'Type_Changes': {Type : 1}, 'Object_Changes': {Object : 1}}
     
    else:
        if Username not in months[month]['Username_Changes']:
            months[month]['Username_Changes'][Username] = 1
        else:
            months[month]['Username_Changes'][Username] += 1
        
        if Type not in months[month]['Type_Changes']:
            months[month]['Type_Changes'][Type] = 1
        else:
            months[month]['Type_Changes'][Type] += 1

        if Object not in months[month]['Object_Changes']:
            months[month]['Object_Changes'][Object] = 1
        else:
            months[month]['Object_Changes'][Object] += 1
            
        months[month]['Total_Changes'] += 1
        
# Sort months
months = dict(sorted(months.items(), key=lambda month: datetime.datetime.strptime(month[0], '%B %Y')))


# In[27]:


# Compartmentalise data in a daily basis

from datetime import datetime

dates = {}
for row in database_df.iterrows():
    date_str = row[1]['DateTime'][:10]
    day = date_str[:10]
    
    Username = row[1]['UserName']
    Type = row[1]['Type']
    Object = row[1]['Object']
    
    if day not in dates:
        dates[day] = {'Total_Changes': 1, 'Username_Changes':{Username: 1}, 'Type_Changes': {Type : 1}, 'Object_Changes': {Object : 1}}
    else:
        if Username not in dates[day]['Username_Changes']:
            dates[day]['Username_Changes'][Username] = 1
        else:
            dates[day]['Username_Changes'][Username] += 1
        
        if Type not in dates[day]['Type_Changes']:
            dates[day]['Type_Changes'][Type] = 1
        else:
            dates[day]['Type_Changes'][Type] += 1

        if Object not in dates[day]['Object_Changes']:
            dates[day]['Object_Changes'][Object] = 1
        else:
            dates[day]['Object_Changes'][Object] += 1
            
        dates[day]['Total_Changes'] += 1   
    


# In[28]:


# Write Weekly and Monthly Totals Data to File

import xlsxwriter

workbook = xlsxwriter.Workbook('Weekly and Monthly Data.xlsx')
weekly_totals = workbook.add_worksheet('Weekly Totals Data')

weekly_totals.write("A1", 'Week Starting')
weekly_totals.write("B1", "Total Changes Made")

count = 1

for week in weeks:
    weekly_totals.write(count, 0, str(weeks[week]['Week Starting']))
    weekly_totals.write(count, 1, weeks[week]['Total_Changes'])
    count += 1


    
    
monthly_totals = workbook.add_worksheet('Monthly Totals Data')

monthly_totals.write('A1', 'Month')
monthly_totals.write('B1', 'Total Changes Made')

count = 1

for month in months:
    monthly_totals.write(count, 0, month)
    monthly_totals.write(count, 1, months[month]['Total_Changes'])
    count += 1

    
workbook.close()

    


# In[29]:


# Write weekly sliced data to file

workbook = xlsxwriter.Workbook('Weekly Slicable Data.xlsx')
data = workbook.add_worksheet('Data')

data.write("A1", "Week")
data.write("B1", "Username")
data.write("C1","Changes Made by User")
data.write("D1", "Type")
data.write("E1","Changes made of this Type")
data.write("F1", "Object")
data.write("G1", "Changes made of this object")

count = 1
for week in weeks:
    original_count = count
    maximum = max(len(weeks[week]['Username_Changes']), len(weeks[week]['Type_Changes']), len(weeks[week]['Object_Changes']))
    
    for i in range(maximum):
        data.write(count, 0, str(weeks[week]['Week Starting']))
        count += 1
        
    next_count = count + 1
    
    
    count = original_count
    for user in weeks[week]['Username_Changes']:
        data.write(count,1, user)
        data.write(count, 2, weeks[week]['Username_Changes'][user])
        count += 1
    
    count = original_count
    for Type in weeks[week]['Type_Changes']:
        data.write(count,3,Type)
        data.write(count,4, weeks[week]['Type_Changes'][Type])
        count += 1
    
    count = original_count
    for Object in weeks[week]['Object_Changes']:
        data.write(count, 5, Object)
        data.write(count ,6, weeks[week]['Object_Changes'][Object])
        count += 1
        
    count = next_count
    
    
workbook.close()
        


# In[30]:


# Write monthly sliced data to file.
workbook = xlsxwriter.Workbook('Monthly Slicable Data.xlsx')
data = workbook.add_worksheet('Data')

data.write("A1", "Month")
data.write("B1", "Username")
data.write("C1","Changes Made by User")
data.write("D1", "Type")
data.write("E1","Changes made of this Type")
data.write("F1", "Object")
data.write("G1", "Changes made of this object")

count = 1
for month in months:
    original_count = count
    maximum = max(len(months[month]['Username_Changes']), len(months[month]['Type_Changes']), len(months[month]['Object_Changes']))
    
    for i in range(maximum):
        data.write(count, 0, str(month))
        count += 1
        
    next_count = count + 1
    
    
    count = original_count
    for user in months[month]['Username_Changes']:
        data.write(count,1, user)
        data.write(count, 2, months[month]['Username_Changes'][user])
        count += 1
    
    count = original_count
    for Type in months[month]['Type_Changes']:
        data.write(count,3,Type)
        data.write(count,4, months[month]['Type_Changes'][Type])
        count += 1
    
    count = original_count
    for Object in months[month]['Object_Changes']:
        data.write(count, 5, Object)
        data.write(count ,6, months[month]['Object_Changes'][Object])
        count += 1
        
    count = next_count
    
    
workbook.close()


# In[31]:


# Write daily sliced data to file.


workbook = xlsxwriter.Workbook('Daily Slicable Data.xlsx')
data = workbook.add_worksheet('Data')

data.write("A1", "Date")
data.write("B1", "Username")
data.write("C1","Changes Made by User")
data.write("D1", "Type")
data.write("E1","Changes made of this Type")
data.write("F1", "Object")
data.write("G1", "Changes made of this object")

count = 1
for day in dates:
    original_count = count
    maximum = max(len(dates[day]['Username_Changes']), len(dates[day]['Type_Changes']), len(dates[day]['Object_Changes']))
    
    for i in range(maximum):
        data.write(count, 0, str(day))
        count += 1
        
    next_count = count + 1
    
    
    count = original_count
    for user in dates[day]['Username_Changes']:
        data.write(count,1, user)
        data.write(count, 2, dates[day]['Username_Changes'][user])
        count += 1
    
    count = original_count
    for Type in dates[day]['Type_Changes']:
        data.write(count,3,Type)
        data.write(count,4, dates[day]['Type_Changes'][Type])
        count += 1
    
    count = original_count
    for Object in dates[day]['Object_Changes']:
        data.write(count, 5, Object)
        data.write(count ,6, dates[day]['Object_Changes'][Object])
        count += 1
        
    count = next_count
    
    
workbook.close()


# In[32]:


# Write Daily Totals Data to File

import xlsxwriter

workbook = xlsxwriter.Workbook('Daily Totals Data.xlsx')
data = workbook.add_worksheet('Data')

data.write("A1", 'Date')
data.write("B1", "Total Changes Made")

count = 1

for day in dates:
    data.write(count, 0, str(day))
    data.write(count, 1, dates[day]['Total_Changes'])
    count += 1

    
workbook.close()


# In[33]:


# User summary tool, with daily slice incorporated.

from datetime import datetime

users = {}
for row in database_df.iterrows():
    date_str = row[1]['DateTime'][:10]
    day = date_str[:10]
    
    user = row[1]['UserName']
    Type = row[1]['Type']
    Object = row[1]['Object']
    
    if user not in users:
        users[user] = {'Total_Changes': 1, 'dates' : {day : {'Type_Changes': {Type : 1}, 'Object_Changes': {Object : 1}}}}
    else:
        if day not in users[user]['dates']:
            users[user]['dates'][day]= {'Type_Changes': {Type : 1}, 'Object_Changes': {Object : 1}}
            
        else:

            if Type not in users[user]['dates'][day]['Type_Changes']:
                users[user]['dates'][day]['Type_Changes'][Type] = 1
            else:
                users[user]['dates'][day]['Type_Changes'][Type] += 1

            if Object not in users[user]['dates'][day]['Object_Changes']:
                users[user]['dates'][day]['Object_Changes'][Object] = 1
            else:
                users[user]['dates'][day]['Object_Changes'][Object] += 1

    users[user]['Total_Changes'] += 1
        


# In[34]:


# Write user tool data to file.
import xlsxwriter

workbook = xlsxwriter.Workbook('User Tool Data.xlsx')
data = workbook.add_worksheet('Data')

data.write("A1", 'User')
data.write("B1", "Date")
data.write("C1", "Type")
data.write("D1", "Type Changes Made")
data.write("E1", "Object")
data.write("F1", "Object Changes Made")
data.write("G1", "Total Changes Made")




count = 1
for user in users:
    
    for day in users[user]['dates']:
        total_changes = 0
        original_count = count
        maximum = max(len(users[user]['dates'][day]['Type_Changes']), len(users[user]['dates'][day]['Object_Changes']))

        for i in range(maximum):
            data.write(count, 1, str(day))
            data.write(count, 0, user )
            count += 1

        next_count = count + 1

        count = original_count
        for Type in users[user]['dates'][day]['Type_Changes']:
            data.write(count,2,Type)
            data.write(count,3, users[user]['dates'][day]['Type_Changes'][Type])
            total_changes += users[user]['dates'][day]['Type_Changes'][Type]
            count += 1

        count = original_count
        for Object in users[user]['dates'][day]['Object_Changes']:
            data.write(count, 4, Object)
            data.write(count ,5, users[user]['dates'][day]['Object_Changes'][Object])
            count += 1
            
        count = original_count
        for i in range(maximum):
            data.write(count, 6, total_changes)
            count += 1
            break
        
        

        count = next_count

workbook.close()

