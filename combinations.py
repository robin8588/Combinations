import itertools
import pandas as pd
from datetime import datetime


class CombinationMaker:
    def __init__(self, column, pick, sheets):
        self.columns = column
        self.pick = pick
        self.sheets = sheets

    def GenExcel(self, path):
        lists = itertools.combinations(self.columns, self.pick)
        row_lists = []
        for row in lists:
            dit = {}
            for i in row:
                dit[i] = i
            row_lists.append(dit)
        df = pd.DataFrame(row_lists, columns=self.columns)
        df.index = df.index+1
        rows = len(row_lists)
        print('total rows: ', rows)
        writer = pd.ExcelWriter(path)
        for i in range(self.sheets):
            min = int(rows/self.sheets) * i
            max = int(rows/self.sheets) * (i+1)
            if i == self.sheets-1:
                max = rows
            sheet_name = str(min+1)+'-'+str(max)
            print(sheet_name,datetime.now())
            dfpart = df[(df.index <= max) & (df.index > min)]
            dfpart.to_excel(writer, sheet_name=sheet_name)
            writer.save()

    def GenLast(self, path):
        lists = itertools.combinations(self.columns, self.pick)
        row_lists = []
        for row in lists:
            dit = {}
            for i in row:
                dit[i] = i
            row_lists.append(dit)
        df = pd.DataFrame(row_lists, columns=self.columns)
        df.index = df.index+1
        rows = len(row_lists)
        print('total rows: ', rows)
        writer = pd.ExcelWriter(path)
        for i in range(self.sheets):
            if i == 3:
                print('start: ',i,datetime.now())
                min = int(rows/self.sheets) * i
                max = int(rows/self.sheets) * (i+1)
                if i == self.sheets-1:
                    max = rows
                sheet_name = str(min+1)+'-'+str(max)
                print(sheet_name)
                dfpart = df[(df.index <= max) & (df.index > min)]
                dfpart.to_excel(writer, sheet_name=sheet_name)
                writer.save()
            else:
                print('skip: ',i)
