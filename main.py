from datetime import datetime
import combinations

path = "combinations.xlsx"
columns = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33']
pick = 6

maker = combinations.CombinationMaker(columns,pick)
print(datetime.now())
maker.GenExcel(path)
print(datetime.now())