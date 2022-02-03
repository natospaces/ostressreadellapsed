import pandas as np
import os
from pathlib import Path
import xlsxwriter
workbook = xlsxwriter.Workbook('outfiles.xlsx')
worksheet = workbook.add_worksheet()


def get_elapsed(location):
#location = 'n50\spl_module_data_tvp31287_1.out'


    with open(location,'r') as f:
        listl=[]
        for line in f:
            strip_lines=line.strip()
            listli=strip_lines.split('\x00')
            m=listl.append(listli)

        return(''.join(listl[len(listl)-3]))


folder = Path(r'E:\readOUT\n102')
elapsed_list = []   
for file in os.listdir(folder):
     filename = os.fsdecode(file)
     if filename.endswith(".out"): 
         print(get_elapsed(os.path.join(folder, filename)))
         elapsed_time = get_elapsed(os.path.join(folder, filename))
         #get_elapsed(os.path.join(folder, filename))
         eq_index = elapsed_time.find("=")
         ms_index = elapsed_time.find("ms")
         print(elapsed_time[eq_index + 2:ms_index])
         eq_index2 = elapsed_time.find("=",eq_index + 1)
         ms_index2 = elapsed_time.find("ms",ms_index + 1)
         
         elapsed = elapsed_time[eq_index2 + 2:ms_index2]
         print(type(elapsed))
         elapsed_list.append(elapsed.strip())
         continue
     else:
         continue
row = 0
column = 0
print(elapsed_list)
# iterating through content list
for item in elapsed_list :
 
    # write operation perform
    worksheet.write(row, column, item)
 
    # incrementing the value of row by one
    # with each iterations.
    row += 1
     
workbook.close()