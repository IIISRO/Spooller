import xlwings  as xw
from datetime import datetime
import re


path = 'test.xlsm'
app = xw.App(visible = False)


date_cord = "C17"
# unique_word = 'Inbound Processing'

# cell_with_word = None
# for cell in ws.used_range:
#     if cell.value == unique_word:
#         cell_with_word = cell
#         break
# print(cell_with_word)


ite = None
today = datetime.now()



try:
    wb = app.books.open(path)
    ws = wb.sheets["Spool Puller"]

    while ws.range(date_cord).value:
        start_date_str, end_date_str = ws.range(date_cord).value.replace(' ', '').split("-")
        start_date = datetime.strptime(start_date_str, "%m/%d")
        end_date = datetime.strptime(end_date_str, "%m/%d")

        if start_date.replace(year=today.year) <= today <= end_date.replace(year=today.year):
            ite = ws.range("A" + re.search(r'\d+', date_cord).group()).value
            break

        date_cord = "C" + str(int(re.search(r'\d+', date_cord).group())+1)

    ws.range("B8").value = ite
    wb.macro("GetITEDates")()
    wb.save()
    wb.close()


finally:
    app.quit()