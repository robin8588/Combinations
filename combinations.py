import itertools
import pandas as pd

class CombinationMaker:
    def __init__(self,column,pick):
        self.columns = column
        self.pick = pick
        self.rows = 553784

    def GenExcel(self,path):
        lists = itertools.combinations(self.columns,self.pick)
        row_lists = []
        for row in lists:
            dit = {}
            for i in row:
                dit[i] = i
            row_lists.append(dit)
        df = pd.DataFrame(row_lists, columns=self.columns)
        df.index = df.index+1
        dfup = df[df.index <= 100]
        dfdown = df[df.index > 1107500]
        
        with pd.ExcelWriter(path) as writer:
            dfup.to_excel(writer,sheet_name = "上")
            dfdown.to_excel(writer,sheet_name = "下")
