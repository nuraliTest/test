import pandas as pd
from pymongo import MongoClient
from dotenv import load_dotenv, find_dotenv
import os
from pyexcelerate import Workbook
from datetime import datetime, date, time

# Connect to MONGODB
load_dotenv(find_dotenv())
password = os.environ.get("MONGODB_PWD")
connection_string = f"mongodb+srv://newreality:{password}@cluster0.jc2svp8.mongodb.net/test?retryWrites=true&w=majority"
client = MongoClient(connection_string)

new_db = client["mydatabase"]

class Employees:
    def __init__(self):
        data = {
                'Name': ['Alex', 'Justin', 'Set', 'Carlos', 'Gareth', 'John', 'Bob'],
                'Surname': ['Smur', 'Forman', 'Carey', 'Carey', 'Chapman', 'James', 'James'],
                'Age': [21, 25, 35, 40, 19, 27, 25],
                'Job': ['Python Developer', 'Java Developer', 'Project Manager', 'Enterprise Architect', 
                        'Python Developer', 'IOS Developer', 'Python Developer'],
                'Datetime': [datetime(2022, 1, 1, 9, 45, 12).isoformat(), datetime(2022, 1, 1, 11, 50, 25).isoformat(), 
                            datetime(2022, 1, 1, 10, 0, 45).isoformat(), datetime(2022, 1, 1, 9, 7, 36).isoformat(), 
                            datetime(2022, 1, 1, 11, 54, 10).isoformat(), datetime(2022, 1, 1, 9, 56, 40).isoformat(), 
                            datetime(2022, 1, 1, 9, 52, 45).isoformat()]
                }
        self.df = pd.DataFrame(data)

    def FirstCond(self):
        self.dfcopy1 = self.df[:].copy()
        # Add new column to dataframe
        newCol = [None] * len(self.dfcopy1)
        self.dfcopy1['TimeToEnter'] = newCol
        
        # Change entry time for Developers
        job = self.dfcopy1[self.dfcopy1['Job'].str.contains('Developer')]
        newdf18 = job[(job['Age'] > 18) & (job['Age'] <= 21)] 
        others = job[(job['Age'] < 18) | (job['Age'] > 21)]
        newdf18['TimeToEnter'] = time(9, 0, 0).isoformat()
        others['TimeToEnter'] = time(9, 15, 0).isoformat()
        self.dfcopy1.update(newdf18)
        self.dfcopy1.update(others)
        # print(self.dfcopy1)

        # Save dataframe as Excel file
        self.saveToExcel(self.dfcopy1, "1stDataFrame.xlsx")

        # Insert dataframe to collection
        self.dfToMongo(self.dfcopy1, "18MoreAnd21andLess", new_db)
    

    def SecondCond(self):
        self.dfcopy2 = self.df[:].copy()
        # Add new column to dataframe
        newCol = [None] * len(self.dfcopy2)
        self.dfcopy2['TimeToEnter'] = newCol

        # Change entry time for Employees
        second = self.dfcopy2[self.dfcopy2["Job"].str.contains('Developer|Manager')==False]
        more35 = second[(second['Age'] >= 35)]
        more35['TimeToEnter'] = time(11, 0, 0).isoformat()

        othersDM = self.dfcopy2[self.dfcopy2["Job"].str.contains('Developer|Manager')==True]
        others2 = second[(second['Age'] < 35)]
        others2['TimeToEnter'] = time(11, 30, 0).isoformat()
        othersDM['TimeToEnter'] =time(11, 30, 0).isoformat()
        self.dfcopy2.update(more35)
        self.dfcopy2.update(others2)
        self.dfcopy2.update(othersDM)

        # Save dataframe as Excel file
        self.saveToExcel(self.dfcopy2, "2ndDataFrame.xlsx")

        # Insert dataframe to collection
        self.dfToMongo(self.dfcopy2, "35AndMore", new_db)


    def ThirdCond(self):
        self.dfcopy3 = self.df[:].copy()
        # Add new column to dataframe
        newCol = [None] * len(self.dfcopy3)
        self.dfcopy3['TimeToEnter'] = newCol

        # Change entry time for Employees
        third = self.dfcopy3[self.dfcopy3["Job"].str.contains('Architect')]
        otherss = self.dfcopy3[self.dfcopy3["Job"].str.contains('Architect')==False]
        third['TimeToEnter'] = time(10, 30, 0).isoformat()
        otherss['TimeToEnter'] = time(10, 40, 0).isoformat()
        self.dfcopy3.update(third)
        self.dfcopy3.update(otherss)

        # Save dataframe as Excel file
        self.saveToExcel(self.dfcopy3, "3rdDataFrame.xlsx")
        
        # Insert dataframe to collection
        self.dfToMongo(self.dfcopy3, "ArchitectEnterTime", new_db)

    def saveToExcel(self, dtFr, nameExcel):
        values = [dtFr.columns] + list(dtFr.values)
        wb = Workbook()
        wb.new_sheet('Sheet1', data=values)
        wb.save(nameExcel)

    def dfToMongo(self, datafr, collection, db):
        mycol = db[collection]
        datafr.reset_index(inplace=True)
        data_dict = datafr.to_dict("records")
        # Insert collection
        mycol.insert_many(data_dict)


x = Employees()
x.FirstCond()
x.SecondCond()
x.ThirdCond()




