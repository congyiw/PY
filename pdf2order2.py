# importing required modules 
import pdfplumber
import pandas as pd
import xlwings as xw
import os
import inflect
import datetime 
from openpyxl import load_workbook

cwd = os.getcwd()
pdf = pdfplumber.open(cwd +'\\1.pdf')
pd.options.mode.chained_assignment = None 

df_combined = pd.DataFrame(columns=["Supplier Code", "Qty", "Unit Cost"])

with pdfplumber.open(cwd + '\\1.pdf') as pdf:
    for page in pdf.pages:
        table = page.extract_table()
        if table:
            df_page = pd.DataFrame(table[1:], columns=table[0])
            df_page_subset = df_page[["Supplier Code", "Qty", "Unit Cost"]]
            df_page_subset.dropna(how='all', inplace=True)  
            df_combined = pd.concat([df_combined, df_page_subset])

df_combined.reset_index(drop=True, inplace=True)
print(df_combined.to_string(index=False, header=False))
rows = df_combined.shape[0]
order_total = []

for idx,data in df_combined.iterrows():
    order_row= []
    order_row.append([data[0],data[1],data[2]])# data[0]:'supplier code',data[1]:'qty',data[2]:'unit price'
    order_total.append(order_row)
today_date = datetime.datetime.now().strftime('%d/%m/%Y')

app = xw.App(visible=True,add_book=False)
app.display_alerts=False
wb = app.books.open('order_template.xls')
sht=wb.sheets('Sheet')
sht.range('I2').value = today_date
sht.range('J11').value = f"Project NO:" 
sht.range('I11').value = f"152426" 
sht.range('I7').value = f"B2409000313" 
first_cell_row = 17 #start from row 17
i = 0
for item in order_total:
    s_code = item[0][0].replace("−", "-") if isinstance(item[0][0], str) else item[0][0]
    qty = item[0][1]
    unit_price = item[0][2]
    if s_code:  
        row = first_cell_row + i
        item_number_cell = 'A' + str(row) 
        code_cell = 'B'+ str(row)  #column B
        qty_cell = 'F' + str(row)  #column F
        unit_price_cell = 'H' + str(row)  #column H       
        sht.range(item_number_cell).value = i+1
        sht.range(code_cell).value = s_code
        sht.range(qty_cell).value = qty
        sht.range(unit_price_cell).value = unit_price     
        i += 1
last_written_row = first_cell_row + i - 1

aud_total_value = sht.range('I85').value
p = inflect.engine()
aud_total_text = p.number_to_words(aud_total_value).upper()
output_string = f"Say AUD {aud_total_text} only."
sht.range('B86').value = output_string

# Add the check for last_written_row and remove rows if necessary
if last_written_row < 84:
    # Remove all rows between last_written_row and 84
    sht.range(f'{last_written_row + 1}:{84}').api.Delete()

# Assuming the VLOOKUP formulas are in rows between first_cell_row and last_written_row
sht.range(f'J{first_cell_row}:J{last_written_row}').value = sht.range(f'J{first_cell_row}:J{last_written_row}').value
# Delete columns K and L
sht.range('K:L').api.Delete()
wb.save('B2409000313.xls')
wb.close
app.quit()