from datetime import datetime
import combinations

path = "combinations30.xlsx"
columns = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14',
           '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30']
pick = 7
sheets = 4

maker = combinations.CombinationMaker(columns, pick, sheets)
print(datetime.now())
maker.GenExcel(path)
print(datetime.now())