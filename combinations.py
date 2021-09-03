import itertools
import pandas as pd
from datetime import datetime


class CombinationMaker:
    def __init__(self, column, pick, sheets, path):
        self.columns = column
        self.pick = pick
        self.sheets = sheets
        self.path = path
        self.rows = 0

    def GetList(self):
        print('get table: ', datetime.now())
        lists = itertools.combinations(self.columns, self.pick)
        row_lists = []
        for row in lists:
            dit = {}
            for i in row:
                dit[i] = i
            row_lists.append(dit)
        df = pd.DataFrame(row_lists, columns=self.columns)
        df.index = df.index+1
        self.rows = len(row_lists)
        print('set table: ', datetime.now())
        return df

    def GenSheet(self, df, i):
        file = self.path+'file'+str(i)+'.xlsx'
        print('create: ',file)
        writer = pd.ExcelWriter(file)
        min = int(self.rows/self.sheets) * i
        max = int(self.rows/self.sheets) * (i+1)
        if i == self.sheets-1:
            max = self.rows
        sheet_name = str(min+1)+'-'+str(max)
        print(sheet_name, datetime.now())
        dfpart = df[(df.index <= max) & (df.index > min)]
        dfpart.to_excel(writer, sheet_name=sheet_name)
        writer.save()

    def GenExcel(self, df):
        print('total rows: ', self.rows)
        for i in range(self.sheets):
            self.GenSheet(df, i)