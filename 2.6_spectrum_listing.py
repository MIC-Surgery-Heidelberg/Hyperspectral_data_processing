"""
@author: Alexander Studier-Fischer, Jan Odenthal, Berkin Özdemir, University of Heidelberg
"""
###
# Finding spectrum file in picture Folder, converting and appending it to Excel File


import os
import pandas as pd
from pathlib import Path
import numpy as np
import xlsxwriter
import glob
from datetime import datetime

###

specify_folder = "/data" # "/data"

folder_in_use = '/_hypergui_1'

label = "_labelling_001"

use_label = True

start_date = ""
end_date = ""                       ### If new end date wanted, change here, make sure to use the same format

###

if start_date  != "" or end_date != "":
    tag = "_specified"
else:
    tag = ""
if start_date == "":
    start_date = "2019_12_14_00_00_01"
if end_date == "":
    end_date = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')


all_paths = list()
if use_label == True:
    tag = "_specified"
    for i in glob.glob('./data/*/2*/' + label + '*.txt'):
        if os.path.basename(os.path.abspath(Path(i).parent)) > start_date and os.path.basename(os.path.abspath(Path(i).parent)) < end_date:
            all_paths.append(os.path.abspath(os.path.abspath(Path(i).parent)))
elif use_label == False:
    for i in glob.glob('./data/*/2*'):
        print(start_date)
        print(end_date)
        if os.path.basename(i) > start_date and os.path.basename(i) < end_date:
            all_paths.append(os.path.abspath(i))

# Main directory
home = os.getcwd().replace("\\", "/")
# Excel Set-up
writer = pd.ExcelWriter(home + '/spectrum_listing_results' + folder_in_use[1:] + tag + '.xlsx', engine= 'xlsxwriter') # Has to be changed individually
home = home + specify_folder
workbook = writer.book
cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
cell_format.set_bold()
cell_format1 = workbook.add_format({'bold': True})
data_path = home  #Has to be changed individually
data_folders = os.listdir(data_path)
row_reflectance_0 = 3
row_absorbance_0 = 3
row_reflectance_1 = 3
row_absorbance_1 = 3
row_reflectance_2 = 3
row_absorbance_2 = 3
row_op_ref_0 = 2
col_op_ref_0 = 3
row_op_ab_0 = 2
col_op_ab_0 = 3
row_op_ref_1 = 2
col_op_ref_1 = 3
row_op_ab_1 = 2
col_op_ab_1 = 3
row_op_ref_2 = 2
col_op_ref_2 = 3
row_op_ab_2 = 2
col_op_ab_2 = 3

# Reflectance_list Sheet 0
    
print("sheet1")

used_names = []
this_round = ''
reflectance_list_glob = []
for i in all_paths:
    for ii in glob.glob(i + '/' + folder_in_use + '/spectrum_fromCSV1*data.xlsx', recursive=True):
        reflectance_list_glob.append(ii)
reflectance_list_glob = list(set(reflectance_list_glob))

for filename in sorted(reflectance_list_glob):
    print(filename)
    base = pd.DataFrame(pd.read_excel(filename,'0_derivative', decimal=',', header=None, usecols=[0]))
    converted = pd.DataFrame(pd.read_excel(filename, '0_derivative', decimal= ',', header=None, usecols=[1]))
    converted.to_excel(writer, sheet_name='Reflectance_list_0_derivative', header=False, index=False, startrow=5, startcol= row_reflectance_0)
    base.to_excel(writer, sheet_name='Reflectance_list_0_derivative', header=False, index=False, startrow=5, startcol=1,)
    row_reflectance_0 += 1

    image_dir_ref = Path(filename).parent.parent.absolute()
    op_dir_ref = Path(image_dir_ref).parent.absolute()
    image_dir_ref = os.path.basename(image_dir_ref)
    
    op_dir_ref = os.path.basename(op_dir_ref)

    worksheet1 = writer.sheets['Reflectance_list_0_derivative']
    if op_dir_ref != this_round:
        worksheet1.write_string(row_op_ref_0, col_op_ref_0, op_dir_ref)
    if any([name == image_dir_ref for name in used_names]):
        cell_format = workbook.add_format()
        cell_format.set_font_color('red')
        worksheet1.write_string(row_op_ref_0 + 1, col_op_ref_0, image_dir_ref, cell_format)
    else:
        worksheet1.write_string(row_op_ref_0 + 1, col_op_ref_0, image_dir_ref)
    used_names.append(image_dir_ref)
    col_op_ref_0 += 1
    this_round = op_dir_ref

# Absorbance_list Sheet 0
print("sheet2")
this_round = ''
absorbance_list_glob = []
for i in all_paths:
    for ii in glob.glob(i+ '/**/' + folder_in_use + '/spectrum_fromCSV5*data.xlsx', recursive = True):
        absorbance_list_glob.append(ii)
absorbance_list_glob = list(set(absorbance_list_glob))
for filename in sorted(absorbance_list_glob):
    print(filename)
    base = pd.DataFrame(pd.read_excel(filename,'0_derivative', decimal=',', header=None, usecols=[0]))
    converted = pd.DataFrame(pd.read_excel(filename, '0_derivative',  decimal= ',', header=None, usecols=[1]))
    converted.to_excel(writer, sheet_name='Absorbance_list_0_derivative', header=False, index=False, startrow=5, startcol= row_absorbance_0)
    base.to_excel(writer, sheet_name='Absorbance_list_0_derivative', header=False, index=False, startrow=5, startcol=1,)
    row_absorbance_0 += 1

    image_dir_ab = Path(filename).parent.parent.absolute()
    op_dir_ab = Path(image_dir_ab).parent.absolute()
    image_dir_ab = os.path.basename(image_dir_ab)
    op_dir_ab = os.path.basename(op_dir_ab)

    worksheet2 = writer.sheets['Absorbance_list_0_derivative']
    if op_dir_ab != this_round:
        worksheet2.write_string(row_op_ab_0, col_op_ab_0, op_dir_ab)
    worksheet2.write_string(row_op_ab_0 + 1, col_op_ab_0, image_dir_ab)
    col_op_ab_0 += 1
    this_round = op_dir_ab
    
# Reflectance_list Sheet 1
print("sheet3")
this_round = ''
reflectance_list_glob = list(set(reflectance_list_glob))
for filename in sorted(reflectance_list_glob):
    print(filename)
    base = pd.DataFrame(pd.read_excel(filename,'1_derivative', decimal=',', header=None, usecols=[0]))
    converted = pd.DataFrame(pd.read_excel(filename, '1_derivative', decimal= ',', header=None, usecols=[1]))
    converted.to_excel(writer, sheet_name='Reflectance_list_1_derivative', header=False, index=False, startrow=5, startcol= row_reflectance_1)
    base.to_excel(writer, sheet_name='Reflectance_list_1_derivative', header=False, index=False, startrow=5, startcol=1,)
    row_reflectance_1 += 1

    image_dir_ref = Path(filename).parent.parent.absolute()
    op_dir_ref = Path(image_dir_ref).parent.absolute()
    image_dir_ref = os.path.basename(image_dir_ref)
    op_dir_ref = os.path.basename(op_dir_ref)

    worksheet3 = writer.sheets['Reflectance_list_1_derivative']
    if op_dir_ref != this_round:
        worksheet3.write_string(row_op_ref_1, col_op_ref_1, op_dir_ref)
    worksheet3.write_string(row_op_ref_1 + 1, col_op_ref_1, image_dir_ref)
    col_op_ref_1 += 1
    this_round = op_dir_ref

# Absorbance_list Sheet 1
print("sheet4")
this_round = ''
absorbance_list_glob = list(set(absorbance_list_glob))
for filename in sorted(absorbance_list_glob):
    print(filename)
    base = pd.DataFrame(pd.read_excel(filename,'1_derivative', decimal=',', header=None, usecols=[0]))
    converted = pd.DataFrame(pd.read_excel(filename, '1_derivative', decimal= ',', header=None, usecols=[1]))
    converted.to_excel(writer, sheet_name='Absorbance_list_1_derivative', header=False, index=False, startrow=5, startcol= row_absorbance_1)
    base.to_excel(writer, sheet_name='Absorbance_list_1_derivative', header=False, index=False, startrow=5, startcol=1,)
    row_absorbance_1 += 1

    image_dir_ab = Path(filename).parent.parent.absolute()
    op_dir_ab = Path(image_dir_ab).parent.absolute()
    image_dir_ab = os.path.basename(image_dir_ab)
    op_dir_ab = os.path.basename(op_dir_ab)

    worksheet4 = writer.sheets['Absorbance_list_1_derivative']
    if op_dir_ab != this_round:
        worksheet4.write_string(row_op_ab_1, col_op_ab_1, op_dir_ab)
    worksheet4.write_string(row_op_ab_1 + 1, col_op_ab_1, image_dir_ab)
    col_op_ab_1 += 1
    this_round = op_dir_ab
    
# Reflectance_list Sheet 2
print("sheet5")
this_round = ''
reflectance_list_glob = list(set(reflectance_list_glob))
for filename in sorted(reflectance_list_glob):
    print(filename)
    base = pd.DataFrame(pd.read_excel(filename,'2_derivative', decimal=',', header=None, usecols=[0]))
    converted = pd.DataFrame(pd.read_excel(filename,'2_derivative', decimal= ',', header=None, usecols=[1]))
    converted.to_excel(writer, sheet_name='Reflectance_list_2_derivative', header=False, index=False, startrow=5, startcol= row_reflectance_2)
    base.to_excel(writer, sheet_name='Reflectance_list_2_derivative', header=False, index=False, startrow=5, startcol=1,)
    row_reflectance_2 += 1

    image_dir_ref = Path(filename).parent.parent.absolute()
    op_dir_ref = Path(image_dir_ref).parent.absolute()
    image_dir_ref = os.path.basename(image_dir_ref)
    op_dir_ref = os.path.basename(op_dir_ref)

    worksheet5 = writer.sheets['Reflectance_list_2_derivative']
    if op_dir_ref != this_round:
        worksheet5.write_string(row_op_ref_2, col_op_ref_2, op_dir_ref)
    worksheet5.write_string(row_op_ref_2 + 1, col_op_ref_2, image_dir_ref)
    col_op_ref_2 += 1
    this_round = op_dir_ref

# Absorbance_list Sheet
print("sheet6")
this_round = ''
absorbance_list_glob = list(set(absorbance_list_glob))
for filename in sorted(absorbance_list_glob):
    print(filename)
    base = pd.DataFrame(pd.read_excel(filename,'2_derivative', decimal=',', header=None, usecols=[0]))
    converted = pd.DataFrame(pd.read_excel(filename,'2_derivative', decimal= ',', header=None, usecols=[1]))
    converted.to_excel(writer, sheet_name='Absorbance_list_2_derivative', header=False, index=False, startrow=5, startcol= row_absorbance_2)
    base.to_excel(writer, sheet_name='Absorbance_list_2_derivative', header=False, index=False, startrow=5, startcol=1,)
    row_absorbance_2 += 1

    image_dir_ab = Path(filename).parent.parent.absolute()
    op_dir_ab = Path(image_dir_ab).parent.absolute()
    image_dir_ab = os.path.basename(image_dir_ab)
    op_dir_ab = os.path.basename(op_dir_ab)

    worksheet6 = writer.sheets['Absorbance_list_2_derivative']
    if op_dir_ab != this_round:
        worksheet6.write_string(row_op_ab_2, col_op_ab_2, op_dir_ab)
    worksheet6.write_string(row_op_ab_2 + 1, col_op_ab_2, image_dir_ab)
    col_op_ab_2 += 1
    this_round = op_dir_ab

# Reflectance_Sheet Set-up
worksheet1.write(0, 3, 'Reflectance (CSV1)', cell_format)
worksheet1.set_row(2, 15, cell_format1)
worksheet1.set_row(3, 15, cell_format1)
worksheet1.set_column(0, 0, 3)
worksheet1.set_column(1, 1, 10)
worksheet1.set_column(2, 2, 3)
worksheet1.freeze_panes(4, 3)
worksheet1.write(3, 1, '=COUNTA(D4:AAA4)')
worksheet1.write(2, 1, '=COUNTA(D3:AAA3)')

# Absorbance Sheet Set-up
worksheet2.write(0, 3, 'Absorbance (CSV5)', cell_format)
worksheet2.set_row(2, 15, cell_format1)
worksheet2.set_row(3, 15, cell_format1)
worksheet2.set_column(0, 0, 3)
worksheet2.set_column(1, 1, 10)
worksheet2.set_column(2, 2, 3)
worksheet2.freeze_panes(4, 3)
worksheet2.write(3, 1, '=COUNTA(D4:AAA4)')
worksheet2.write(2, 1, '=COUNTA(D3:AAA3)')

# Reflectance_Sheet Set-up
worksheet3.write(0, 3, 'Reflectance (CSV1)', cell_format)
worksheet3.set_row(2, 15, cell_format1)
worksheet3.set_row(3, 15, cell_format1)
worksheet3.set_column(0, 0, 3)
worksheet3.set_column(1, 1, 10)
worksheet3.set_column(2, 2, 3)
worksheet3.freeze_panes(4, 3)
worksheet3.write(3, 1, '=COUNTA(D4:AAA4)')
worksheet3.write(2, 1, '=COUNTA(D3:AAA3)')

# Absorbance Sheet Set-up
worksheet4.write(0, 3, 'Absorbance (CSV5)', cell_format)
worksheet4.set_row(2, 15, cell_format1)
worksheet4.set_row(3, 15, cell_format1)
worksheet4.set_column(0, 0, 3)
worksheet4.set_column(1, 1, 10)
worksheet4.set_column(2, 2, 3)
worksheet4.freeze_panes(4, 3)
worksheet4.write(3, 1, '=COUNTA(D4:AAA4)')
worksheet4.write(2, 1, '=COUNTA(D3:AAA3)')

# Reflectance_Sheet Set-up
worksheet5.write(0, 3, 'Reflectance (CSV1)', cell_format)
worksheet5.set_row(2, 15, cell_format1)
worksheet5.set_row(3, 15, cell_format1)
worksheet5.set_column(0, 0, 3)
worksheet5.set_column(1, 1, 10)
worksheet5.set_column(2, 2, 3)
worksheet5.freeze_panes(4, 3)
worksheet5.write(3, 1, '=COUNTA(D4:AAA4)')
worksheet5.write(2, 1, '=COUNTA(D3:AAA3)')

# Absorbance Sheet Set-up
worksheet6.write(0, 3, 'Absorbance (CSV5)', cell_format)
worksheet6.set_row(2, 15, cell_format1)
worksheet6.set_row(3, 15, cell_format1)
worksheet6.set_column(0, 0, 3)
worksheet6.set_column(1, 1, 10)
worksheet6.set_column(2, 2, 3)
worksheet6.freeze_panes(4, 3)
worksheet6.write(3, 1, '=COUNTA(D4:AAA4)')
worksheet6.write(2, 1, '=COUNTA(D3:AAA3)')

writer.save()