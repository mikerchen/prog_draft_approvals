from datetime import datetime

test_date = '2021/06'

my_date = datetime.strptime(test_date, "%Y/%m")

print(my_date)