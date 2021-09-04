"""
@author: Alexander Studier-Fischer, Jan Odenthal, Berkin Oezdemir, University of Heidelberg
"""
###
# Finding spectrum file in picture Folder, converting and appending it to Excel File


import os
import pandas as pd
from pathlib import Path
import numpy as np
import glob
import matplotlib.pyplot as plt
import shutil
from datetime import datetime

###

tag = ""     #endung hier eingeben

specify_folder = "/data" # "/data"

folder_in_use = '/_hypergui_1'

label = ""

use_label = True

start_date = ""
end_date = ""



###

wl_limits =[500, 995]


blue_line_thickness = 2.0
gray_line_thickness = 0.7
black_line_thickness = 1.7

gray = 30

x_tick_fontsize = 10
y_tick_fontsize = 10

x_axis_thickness = 1.2
y_axis_thickness = 1.2

x_axis_tick_width = 1.2
y_axis_tick_width = 1.2

x_axis_tick_length = 4
y_axis_tick_length = 4

ceil_ref_der_0=1
ceil_abs_der_0=1
ceil_ref_der_1=0.1
ceil_abs_der_1=1
ceil_ref_der_2=0.04
ceil_abs_der_2=1

floor_ref_der_0=0
floor_abs_der_0=0
floor_ref_der_1=-0.035
floor_abs_der_1=0
floor_ref_der_2=-0.025
floor_abs_der_2=0

ceil_ref_der_0_l1=0.022
ceil_abs_der_0_l1=1
ceil_ref_der_1_l1=0.05
ceil_abs_der_1_l1=1
ceil_ref_der_2_l1=0.05
ceil_abs_der_2_l1=1

floor_ref_der_0_l1=0
floor_abs_der_0_l1=0
floor_ref_der_1_l1=-0.035
floor_abs_der_1_l1=0
floor_ref_der_2_l1=-0.025
floor_abs_der_2_l1=0

margin = 0.05

resolution = 300

wl_limits[1] = wl_limits[1] + 5
wl_int = int(((wl_limits[1]-wl_limits[0])/5))


if start_date == "":
    start_date = "2018_01_01_00_00_01"
if end_date == "":
    end_date = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')

all_paths = list()
all_paths_parents = list()
if use_label == True:
    for i in glob.glob('./data/*/2*/' + label + '*.txt'):
        if os.path.basename(Path(i).parent) > start_date and os.path.basename(Path(i).parent) < end_date:
            all_paths.append(os.path.abspath(Path(i).parent))
            if os.path.basename(Path(i).parent.parent) not in all_paths_parents:
                all_paths_parents.append(os.path.basename(Path(i).parent.parent))
elif use_label == False:
    for i in glob.glob('./data/*/2*'):
        if os.path.basename(i) > start_date and os.path.basename(i) < end_date:
            all_paths.append(os.path.abspath(i))
            if os.path.basename(Path(i).parent) not in all_paths_parents:
                all_paths_parents.append(os.path.basename(Path(i).parent))


if gray < 0:
    gray = 100
if gray >100:
    gray = 100
gray = 1 - gray/100
color_gray = (gray, gray, gray)

# Main directory
home = os.getcwd().replace("\\", "/")
# Excel Set-up
writer = pd.ExcelWriter(home + '/mean_and_sd_extraction_stabw_results' + folder_in_use[1:] + tag + '.xlsx', engine= 'xlsxwriter') # Has to be changed individually
workbook = writer.book
cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
cell_format.set_bold()
cell_format1 = workbook.add_format({'bold': True})
data_path = home + "/data"#Has to be changed individually
data_folders = os.listdir(data_path)


plotfolder = home + '/mean_and_sd_extraction_stabw_plots' + folder_in_use[1:] + tag
if os.path.exists(plotfolder):
    shutil.rmtree(plotfolder, ignore_errors=True)
if not os.path.exists(plotfolder):
    os.mkdir(plotfolder)

baseX = np.arange(wl_limits[0],wl_limits[1],5)

def make_sheet(xlsx_name, sheet_name, sheet_name_2, foldername, baseX, normalize, floor_ref_der, ceil_ref_der):
    if not os.path.exists(plotfolder + '/'+ foldername):
        os.mkdir(plotfolder + '/'+ foldername)
    if not os.path.exists(plotfolder + '/'+ foldername +'_fixed_scale'):
        os.mkdir(plotfolder + '/'+ foldername + '_fixed_scale')
    fig = plt.figure()
    ax = fig.add_subplot(111)
    row_reflectance = 5
    col_op_ref = 5
    row_op_ref = 2
    yMax=0
    yMin=0
    this_round = ''
    total=vals=np.array([], dtype=np.int64).reshape(0,wl_int)
    for folder in all_paths_parents:
        name=folder
        vals=np.array([], dtype=np.int64).reshape(0,wl_int)
        for pic in all_paths:
            if os.path.basename(Path(pic).parent) == folder:
                for filename in glob.glob(pic+ '/' + folder_in_use + '/' + xlsx_name):
                    base = pd.DataFrame(pd.read_excel(filename, sheet_name, decimal=',', header=None, usecols=[0],skiprows=int((wl_limits[0]-500)/5), nrows=wl_int))
                    converted = pd.DataFrame(pd.read_excel(filename, sheet_name, decimal= ',', header=None, usecols=[1],skiprows=int((wl_limits[0]-500)/5), nrows=wl_int))
    
                    image_dir_ref = Path(filename).parent.parent.absolute()
                    op_dir_ref = Path(image_dir_ref).parent.absolute()
                    image_dir_ref = os.path.basename(image_dir_ref)
                    op_dir_ref = os.path.basename(op_dir_ref)
    
                    vals = np.vstack([vals, np.asarray(converted[1])])
                
        if(vals.shape[0]>=1):
            if normalize: 
                vals = vals / np.linalg.norm(vals, ord = 1, axis = 1, keepdims = True)
                vals = np.nan_to_num(vals, copy = False)
            pig_mean=np.mean(vals, axis=0)
            pig_sd=np.std(vals, axis=0,  ddof = 1)
            if np.min(pig_mean-pig_sd) < yMin:
            	yMin = np.min(pig_mean-pig_sd)
            if np.max(pig_mean+pig_sd)>yMax:
            	yMax = np.max(pig_mean+pig_sd)
            print("writing to worksheet...")
            total=np.vstack([total, pig_mean])
            pd.DataFrame(pig_mean).to_excel(writer, sheet_name=sheet_name_2, header=False, index=False, startrow=5, startcol= row_reflectance)
            row_reflectance += 1
            pd.DataFrame(pig_sd).to_excel(writer, sheet_name=sheet_name_2, header=False, index=False, startrow=5, startcol= row_reflectance)
            row_reflectance += 1
            base.to_excel(writer, sheet_name=sheet_name_2, header=False, index=False, startrow=5, startcol=1,)
            worksheet = writer.sheets[sheet_name_2]
            worksheet.write_string(row_op_ref, col_op_ref, op_dir_ref)
            worksheet.write_string(1, col_op_ref, "mean")
            worksheet.write_string(1, col_op_ref+1, "std")
            col_op_ref += 2
            
        else:
            print(folder +"is empty")
    
    abs_marg = max(yMax, abs(yMin))*margin
    yMax=yMax+abs_marg
    yMin=yMin - abs_marg          
    
    for folder in all_paths_parents:
        name=folder
        vals=np.array([], dtype=np.int64).reshape(0,wl_int)
        for pic in all_paths:
            if os.path.basename(Path(pic).parent) == folder:
                for filename in glob.glob(pic+ '/' + folder_in_use + '/' + xlsx_name):
                    base = pd.DataFrame(pd.read_excel(filename, sheet_name, decimal=',', header=None, usecols=[0],skiprows=int((wl_limits[0]-500)/5), nrows=wl_int))
                    converted = pd.DataFrame(pd.read_excel(filename, sheet_name, decimal= ',', header=None, usecols=[1],skiprows=int((wl_limits[0]-500)/5), nrows=wl_int))
                    vals = np.vstack([vals, np.asarray(converted[1])])
        if(vals.shape[0]>=1):
            if normalize: 
                vals = vals / np.linalg.norm(vals, ord = 1, axis = 1, keepdims = True)
                vals = np.nan_to_num(vals, copy = False)
            pig_mean=np.mean(vals, axis=0)
            pig_sd=np.std(vals, axis=0, ddof = 1 )    
            ax.plot(baseX, pig_mean, color = color_gray, linewidth=gray_line_thickness)
            fig_dummy = plt.figure()
            ax_dummy = fig_dummy.add_subplot(111)
            ax_dummy.spines['left'].set_linewidth(y_axis_thickness)
            ax_dummy.spines['bottom'].set_linewidth(x_axis_thickness)
            ax_dummy.xaxis.set_tick_params(width=x_axis_tick_width, length = x_axis_tick_length)
            ax_dummy.yaxis.set_tick_params(width=y_axis_tick_width, length = y_axis_tick_length)
            ax_dummy.plot(baseX, pig_mean, color =  'b', linewidth=blue_line_thickness)
            ax_dummy.plot(baseX, pig_mean+pig_sd, color = color_gray, linewidth=gray_line_thickness)
            ax_dummy.plot(baseX, pig_mean-pig_sd, color = color_gray, linewidth=gray_line_thickness)
            pic_name = plotfolder + '/'+foldername+'/' + name +'.png'
            ax_dummy.set_ylim(yMin, yMax)
            for tick in ax_dummy.xaxis.get_major_ticks():
                    tick.label.set_fontsize(x_tick_fontsize) 
            for tick in ax_dummy.yaxis.get_major_ticks():
                    tick.label.set_fontsize(y_tick_fontsize) 
            fig_dummy.savefig(pic_name, dpi=resolution, bbox_inches='tight')  
            
            pic_name = plotfolder + '/' + foldername +'_fixed_scale/' + name +'_fixed_scale.png'
            ax_dummy.set_ylim(floor_ref_der, ceil_ref_der)
            for tick in ax_dummy.xaxis.get_major_ticks():
                    tick.label.set_fontsize(x_tick_fontsize) 
            for tick in ax_dummy.yaxis.get_major_ticks():
                    tick.label.set_fontsize(y_tick_fontsize) 
            fig_dummy.savefig(pic_name, dpi=resolution, bbox_inches='tight')  
            plt.close()
            
      
            
    totMean=np.mean(total, axis=0)
    totSd=np.std(total, axis=0,  ddof = 1)
    yMin = np.min(total)
    yMax = np.max(total)
    if np.max(totMean+totSd)> yMax:
        yMax = np.max(totMean+totSd)
    if np.min(totMean-totSd)< yMin:
        yMin = np.min(totMean-totSd)        
    abs_marg = max(yMax, abs(yMin))*margin
    yMax=yMax+abs_marg
    yMin=yMin - abs_marg
            
       
    
    ax.plot(baseX, totMean, color = 'b', linewidth=blue_line_thickness)
    ax.plot(baseX, totMean-totSd, color = (0.1, 0.1, 0.1), linewidth=black_line_thickness)
    ax.plot(baseX, totMean+totSd, color = (0.1, 0.1, 0.1), linewidth=black_line_thickness)
    
    pd.DataFrame(totMean).to_excel(writer, sheet_name=sheet_name_2, header=False, index=False, startrow=5, startcol= 3)
    pd.DataFrame(totSd).to_excel(writer, sheet_name=sheet_name_2, header=False, index=False, startrow=5, startcol= 4)
    worksheet.write_string(2, 3, 'all')
    worksheet.write_string(2, 4, 'all')
    worksheet.write_string(1, 3, "mean")
    worksheet.write_string(1, 4, "std")
        
    ax.set_ylim(yMin, yMax)
    for tick in ax.xaxis.get_major_ticks():
        tick.label.set_fontsize(x_tick_fontsize) 
    for tick in ax.yaxis.get_major_ticks():
        tick.label.set_fontsize(y_tick_fontsize) 
    ax.spines['left'].set_linewidth(y_axis_thickness)
    ax.spines['bottom'].set_linewidth(x_axis_thickness)
    ax.xaxis.set_tick_params(width=x_axis_tick_width, length = x_axis_tick_length)
    ax.yaxis.set_tick_params(width=y_axis_tick_width, length = y_axis_tick_length)
    fig.savefig(plotfolder+'/' + foldername + '/_plot_all.png', dpi=resolution, bbox_inches='tight')
    fig.savefig(plotfolder+'/' + foldername + '/_plot_all.pdf', dpi=resolution, bbox_inches='tight')
    
    ax.set_ylim(floor_ref_der, ceil_ref_der)
    for tick in ax.xaxis.get_major_ticks():
        tick.label.set_fontsize(x_tick_fontsize) 
    for tick in ax.yaxis.get_major_ticks():
        tick.label.set_fontsize(y_tick_fontsize) 
    fig.savefig(plotfolder+'/' + foldername +'_fixed_scale/_plot_all_fixed_scale.png', dpi=resolution,bbox_inches='tight')
    fig.savefig(plotfolder+'/' + foldername +'_fixed_scale/_plot_all_fixed_scale.pdf', dpi=resolution,bbox_inches='tight')
    
    return worksheet

print("sheet1")
worksheet1 = make_sheet('spectrum_fromCSV1*data.xlsx', '0_derivative', 'Reflectance_list_0_derivative', 'reflectance_0_derivative', baseX, False, floor_ref_der_0, ceil_ref_der_0)
worksheet1_l1 = make_sheet('spectrum_fromCSV1*data.xlsx', '0_derivative', 'Reflectance_list_0_dv_l1_norm', 'reflectance_0_derivative_l1_normalized', baseX, True, floor_ref_der_0_l1, ceil_ref_der_0_l1)
print("sheet2")
worksheet2 = make_sheet('spectrum_fromCSV5*data.xlsx', '0_derivative', 'Absorbance_list_0_derivative', 'absorbance_0_derivative', baseX, False, floor_abs_der_0, ceil_abs_der_0)
worksheet2_l1 = make_sheet('spectrum_fromCSV5*data.xlsx', '0_derivative', 'Absorbance_list_0_dv_l1_norm', 'absorbance_0_derivative_l1_normalized', baseX, True, floor_abs_der_0_l1, ceil_abs_der_0_l1)
print("sheet3")
worksheet3 = make_sheet('spectrum_fromCSV1*data.xlsx', '1_derivative', 'Reflectance_list_1_derivative', 'reflectance_1_derivative', baseX, False, floor_ref_der_1, ceil_ref_der_1)
worksheet3_l1 = make_sheet('spectrum_fromCSV1*data.xlsx', '1_derivative', 'Reflectance_list_1_dv_l1_norm', 'reflectance_1_derivative_l1_normalized', baseX, True, floor_ref_der_1_l1, ceil_ref_der_1_l1)
print("sheet4")
worksheet4 = make_sheet('spectrum_fromCSV5*data.xlsx', '1_derivative', 'Absorbance_list_1_derivative', 'absorbance_1_derivative', baseX, False, floor_abs_der_1, ceil_abs_der_1)
worksheet4_l1 = make_sheet('spectrum_fromCSV5*data.xlsx', '1_derivative', 'Absorbance_list_1_dv_l1_norm', 'absorbance_1_derivative_l1_normalized', baseX, True, floor_abs_der_1_l1, ceil_abs_der_1_l1)
print("sheet5")
worksheet5 = make_sheet('spectrum_fromCSV1*data.xlsx', '2_derivative', 'Reflectance_list_2_derivative', 'reflectance_2_derivative', baseX, False, floor_ref_der_2, ceil_ref_der_2)
worksheet5_l1 = make_sheet('spectrum_fromCSV1*data.xlsx', '2_derivative', 'Reflectance_list_2_dv_l1_norm', 'reflectance_2_derivative_l1_normalized', baseX, True, floor_ref_der_2_l1, ceil_ref_der_2_l1)
print("sheet6")
worksheet6 = make_sheet('spectrum_fromCSV5*data.xlsx', '2_derivative', 'Absorbance_list_2_derivative', 'absorbance_2_derivative', baseX, False, floor_abs_der_2, ceil_abs_der_2)
worksheet6_l1 = make_sheet('spectrum_fromCSV5*data.xlsx', '2_derivative', 'Absorbance_list_2_dv_l1_norm', 'absorbance_2_derivative_l1_normalized', baseX, True, floor_abs_der_2_l1, ceil_abs_der_2_l1)

# Reflectance_Sheet Set-up
worksheet1.write(0, 3, 'Reflectance (CSV1)', cell_format)
worksheet1.set_row(2, 15, cell_format1)
worksheet1.set_row(3, 15, cell_format1)
worksheet1.set_column(0, 0, 3)
worksheet1.set_column(1, 1, 10)
worksheet1.set_column(2, 2, 3)
worksheet1.freeze_panes(4, 3)
worksheet1.write(2, 1, '=COUNTA(F3:AAA3)')

# Absorbance Sheet Set-up
worksheet2.write(0, 3, 'Absorbance (CSV5)', cell_format)
worksheet2.set_row(2, 15, cell_format1)
worksheet2.set_row(3, 15, cell_format1)
worksheet2.set_column(0, 0, 3)
worksheet2.set_column(1, 1, 10)
worksheet2.set_column(2, 2, 3)
worksheet2.freeze_panes(4, 3)
worksheet2.write(2, 1, '=COUNTA(F3:AAA3)')

# Reflectance_Sheet Set-up
worksheet3.write(0, 3, 'Reflectance (CSV1)', cell_format)
worksheet3.set_row(2, 15, cell_format1)
worksheet3.set_row(3, 15, cell_format1)
worksheet3.set_column(0, 0, 3)
worksheet3.set_column(1, 1, 10)
worksheet3.set_column(2, 2, 3)
worksheet3.freeze_panes(4, 3)
worksheet3.write(2, 1, '=COUNTA(F3:AAA3)')

# Absorbance Sheet Set-up
worksheet4.write(0, 3, 'Absorbance (CSV5)', cell_format)
worksheet4.set_row(2, 15, cell_format1)
worksheet4.set_row(3, 15, cell_format1)
worksheet4.set_column(0, 0, 3)
worksheet4.set_column(1, 1, 10)
worksheet4.set_column(2, 2, 3)
worksheet4.freeze_panes(4, 3)
worksheet4.write(2, 1, '=COUNTA(F3:AAA3)')

# Reflectance_Sheet Set-up
worksheet5.write(0, 3, 'Reflectance (CSV1)', cell_format)
worksheet5.set_row(2, 15, cell_format1)
worksheet5.set_row(3, 15, cell_format1)
worksheet5.set_column(0, 0, 3)
worksheet5.set_column(1, 1, 10)
worksheet5.set_column(2, 2, 3)
worksheet5.freeze_panes(4, 3)
worksheet5.write(2, 1, '=COUNTA(F3:AAA3)')

# Absorbance Sheet Set-up
worksheet6.write(0, 3, 'Absorbance (CSV5)', cell_format)
worksheet6.set_row(2, 15, cell_format1)
worksheet6.set_row(3, 15, cell_format1)
worksheet6.set_column(0, 0, 3)
worksheet6.set_column(1, 1, 10)
worksheet6.set_column(2, 2, 3)
worksheet6.freeze_panes(4, 3)
worksheet6.write(2, 1, '=COUNTA(F3:AAA3)')

# Reflectance_Sheet Set-up
worksheet1_l1.write(0, 3, 'Reflectance (CSV1) L1-Normalized', cell_format)
worksheet1_l1.set_row(2, 15, cell_format1)
worksheet1_l1.set_row(3, 15, cell_format1)
worksheet1_l1.set_column(0, 0, 3)
worksheet1_l1.set_column(1, 1, 10)
worksheet1_l1.set_column(2, 2, 3)
worksheet1_l1.freeze_panes(4, 3)
worksheet1_l1.write(2, 1, '=COUNTA(F3:AAA3)')

# Absorbance Sheet Set-up
worksheet2_l1.write(0, 3, 'Absorbance (CSV5) L1-Normalized', cell_format)
worksheet2_l1.set_row(2, 15, cell_format1)
worksheet2_l1.set_row(3, 15, cell_format1)
worksheet2_l1.set_column(0, 0, 3)
worksheet2_l1.set_column(1, 1, 10)
worksheet2_l1.set_column(2, 2, 3)
worksheet2_l1.freeze_panes(4, 3)
worksheet2_l1.write(2, 1, '=COUNTA(F3:AAA3)')

# Reflectance_Sheet Set-up
worksheet3_l1.write(0, 3, 'Reflectance (CSV1) L1-Normalized', cell_format)
worksheet3_l1.set_row(2, 15, cell_format1)
worksheet3_l1.set_row(3, 15, cell_format1)
worksheet3_l1.set_column(0, 0, 3)
worksheet3_l1.set_column(1, 1, 10)
worksheet3_l1.set_column(2, 2, 3)
worksheet3_l1.freeze_panes(4, 3)
worksheet3_l1.write(2, 1, '=COUNTA(F3:AAA3)')

# Absorbance Sheet Set-up
worksheet4_l1.write(0, 3, 'Absorbance (CSV5) L1-Normalized', cell_format)
worksheet4_l1.set_row(2, 15, cell_format1)
worksheet4_l1.set_row(3, 15, cell_format1)
worksheet4_l1.set_column(0, 0, 3)
worksheet4_l1.set_column(1, 1, 10)
worksheet4_l1.set_column(2, 2, 3)
worksheet4_l1.freeze_panes(4, 3)
worksheet4_l1.write(2, 1, '=COUNTA(F3:AAA3)')

# Reflectance_Sheet Set-up
worksheet5_l1.write(0, 3, 'Reflectance (CSV1) L1-Normalized', cell_format)
worksheet5_l1.set_row(2, 15, cell_format1)
worksheet5_l1.set_row(3, 15, cell_format1)
worksheet5_l1.set_column(0, 0, 3)
worksheet5_l1.set_column(1, 1, 10)
worksheet5_l1.set_column(2, 2, 3)
worksheet5_l1.freeze_panes(4, 3)
worksheet5_l1.write(2, 1, '=COUNTA(F3:AAA3)')

# Absorbance Sheet Set-up
worksheet6_l1.write(0, 3, 'Absorbance (CSV5) L1-Normalized', cell_format)
worksheet6_l1.set_row(2, 15, cell_format1)
worksheet6_l1.set_row(3, 15, cell_format1)
worksheet6_l1.set_column(0, 0, 3)
worksheet6_l1.set_column(1, 1, 10)
worksheet6_l1.set_column(2, 2, 3)
worksheet6_l1.freeze_panes(4, 3)
worksheet6_l1.write(2, 1, '=COUNTA(F3:AAA3)')

writer.save()
