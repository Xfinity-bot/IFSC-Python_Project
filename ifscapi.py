#Importing requests json pymongo time os openpyxl

import requests
import json
import pymongo   #MongoDB Driver
from time import sleep
import os
from openpyxl import Workbook   #Xlsx


#Database Connection 

client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["IFSC_DB"]  #db
dbcollection = db["banks"] #collection

#DELETION QUERY
#myquery = { "IFSC": "--" }
#dbcollection.delete_one(myquery)

#Clear Screen
os.system('cls')

#Taking IFSC code as input from the user
IFSC_Code=input("Enter IFSC Code")

#Query to check the entered IFSC code is present in DB
myquery = { "IFSC": IFSC_Code }

#returns the count 
flag = dbcollection.count_documents(myquery)

#If present the details is retrieved from the data base
if flag:
    print()
    print("Present In DB")
    sleep(2)   
    os.system('cls')

    #fetching bank details from db
    bankdetails=dbcollection.find_one(myquery)
    print('''
     #Bank       : {}
     #Branch     : {}
     #Address    : {}
     #State      : {}
     #Contact    : {}


    '''.format( bankdetails["BANK"],bankdetails["BRANCH"],bankdetails["ADDRESS"],bankdetails["STATE"],bankdetails["CONTACT"]))
    #creating a spreadsheet
    Spreadsheet = Workbook()
    #Inserting data into spreadsheet
    for x,y in bankdetails.items():
        if x!="_id":
            worksheet = Spreadsheet.active
            i=(x,y)  #tupple
            worksheet.append(i)
    Spreadsheet.save(IFSC_Code+".xlsx")
    
#if data is not present in db it is retrived using api
else :
    #Razorpay API 
    #base url with ifsc code
    URL = "https://ifsc.razorpay.com/"
    result = requests.get(URL+IFSC_Code).content

    #converting json to python format 
    jsontopy=json.loads(result)
    if jsontopy!="Not Found":
        jsonform=json.dumps(jsontopy)
        #Adding to DB
        x = dbcollection.insert_one(jsontopy)
        print()

        print("Adding to DB")
        sleep(2) 
        os.system('cls')
        
        print('''
             Bank       : {}
             Branch     : {}
             Address    : {}
             State      : {}
             Contact    : {}

       
                '''.format( jsontopy["BANK"],jsontopy["BRANCH"],jsontopy["ADDRESS"],jsontopy["STATE"],jsontopy["CONTACT"]))
        #Inserting it into spreadsheet
        Spreadsheet = Workbook()
        for x,y in jsontopy.items():
            if x!="_id":
                worksheet = Spreadsheet.active
                i=(x,y)
                worksheet.append(i)
        Spreadsheet.save(IFSC_Code+".xlsx")

    else:
        print("Invalid IFSC Code")

