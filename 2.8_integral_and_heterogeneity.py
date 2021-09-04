"""
@author: Alexander Studier-Fischer, Jan Odenthal, Berkin Oezdemir, University of Heidelberg
"""
import os
import numpy as np
import glob
import pandas as pd
from scipy.integrate import simps
from numpy import trapz
from math import  sqrt
from sklearn.metrics import mean_squared_error
import similaritymeasures
from datetime import datetime


### Mean and stab Results File name ###

results_file = ".xlsx"

### Integral Results File name ###

new_results_file = ".xlsx"

### Change Hypergui file ###
folder_in_use = '/hypergui_1'
### Change spectrum mode: stabw= 0, stabw.n = 1, spectrum listing = 2,
file_mode = 0
###
wl_limits = [500, 995]


if file_mode == 0:
    file_in_use = '/mean_and_sd_extraction_stabw_results_'
    extra = '_'
elif file_mode == 1:
    file_in_use = '/mean_and_sd_extraction_stabw.n_results_'
    extra = '_'
elif file_mode == 2:
    file_in_use = '/spectrum_listing_results_'
    extra = '_total_'


wl_int = int((((wl_limits[1]+5)-wl_limits[0])/5))

file = os.getcwd()


writer = pd.ExcelWriter(file + new_results_file, engine= 'xlsxwriter') # Has to be changed individually
workbook = writer.book
cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
cell_format.set_bold()
cell_format1 = workbook.add_format({'bold': True})



def area(array):
    simpsons = simps(array.reshape(1,wl_int), dx=1)
    trapez = trapz(array.reshape(1,wl_int), dx=1)
    return float(simpsons)

def get_array(path, mode, der, norm):
    if mode == "ref":
        whichVal = "Reflectance_list_"
    elif mode == "ab":
        whichVal= "Absorbance_list_"
    if der == 0:
        if norm:
            whichDer = "0_dv_l1_norm"
        else:
            whichDer = '0_derivative'
    elif der ==1:
        if norm:
            whichDer = "1_dv_l1_norm"
        else:
            whichDer = '1_derivative'
    elif der ==2:
        if norm:
            whichDer = "2_dv_l1_norm"
        else:
            whichDer = '2_derivative'
    total=np.array([], dtype=np.int64).reshape(0,wl_int)
    if len(glob.glob(path + results_file))==0:
        print(results_file + " not found in " + path)
    filename = glob.glob(path + results_file)[0]
    df = np.asarray(pd.DataFrame(pd.read_excel(filename, whichVal + whichDer, header=None, index=None, skiprows=2)))[:, 3:]
    return df


def integrate(array):
    integral_1 =np.array([], dtype=np.int64).reshape(0,wl_int)
    op_names = np.array([], dtype=np.int64).reshape(0,1)
    main_array = np.array([], dtype=np.int64).reshape(0,wl_int)

    if not file_mode == 2:
        for idx,column in enumerate(array.T):
            if not idx == 0 and idx % 2 == 0:
                integral_1 = np.append(integral_1, area(column[3+ int((wl_limits[0]-500)/5):3+ int((wl_limits[0]-500)/5) + wl_int].reshape(wl_int,1)))
                op_names = np.append(op_names, column[0])
                main_array = np.vstack([main_array, column[3+ int((wl_limits[0]-500)/5):3+ int((wl_limits[0]-500)/5) + wl_int]])
    else:
        for idx,column in enumerate(array.T):
                integral_1 = np.append(integral_1, area(column[3:].reshape(wl_int,1)))
                op_names = np.append(op_names, column[1])
                main_array = np.vstack([main_array, column[3+ int((wl_limits[0]-500)/5):3+ int((wl_limits[0]-500)/5) + wl_int]])

    integral_1.reshape(1, integral_1.shape[0])
    op_names.reshape(1, op_names.shape[0])
    integral_1_mean = np.mean(integral_1)
    integral_sd = np.std(integral_1)
    return(op_names, integral_1, integral_1_mean, integral_sd, main_array)


def compare_pairwise(array, operation):
    comp_array = np.zeros([array.shape[0], array.shape[0]])
    for organ_id1 in np.arange(array.shape[0]):
        for organ_id2 in np.arange(array.shape[0]):
            if operation == "MAPE":
                comp_array[organ_id1, organ_id2] = MAPE(array[organ_id1,:], array[organ_id2,:])
            elif operation == "RMSE":
                comp_array[organ_id1, organ_id2] = RMSE(array[organ_id1,:], array[organ_id2,:])
            elif operation == "Frechet":
                comp_array[organ_id1, organ_id2] = Frechet(array[organ_id1,:], array[organ_id2,:])
            elif operation == "rMAPE":
                comp_array[organ_id1, organ_id2] = rMAPE(array[organ_id1,:], array[organ_id2,:])
            elif operation == "SMAPE":
                comp_array[organ_id1, organ_id2] = SMAPE(array[organ_id1,:], array[organ_id2,:])
            elif operation == "MAD":
                comp_array[organ_id1, organ_id2] = MAD(array[organ_id1,:], array[organ_id2,:])
            elif operation == "MSE":
                comp_array[organ_id1, organ_id2] = MSE(array[organ_id1,:], array[organ_id2,:])
            elif operation == "RSSE":
                comp_array[organ_id1, organ_id2] = RSSE(array[organ_id1,:], array[organ_id2,:])
    return comp_array

def RMSE(array1, array2):
    print('RMSE')
    print(datetime.now())
    return sqrt(mean_squared_error(array1, array2))

def MAD(array1, array2):
    print('MAD')
    print(datetime.now())
    return np.mean(np.abs(array1-array2))

def rMAPE(array1, array2):
    print('rMAPE')
    print(datetime.now())
    for array in [array1, array2]:
        if np.where(array == 0)[0].shape == (1,):
            print(array, 'Array contains 0 Value at line', np.where(array == 0)[0])
            array[np.where(array == 0)[0]] = (array[np.where(array == 0)[0] - 1] + array[np.where(array == 0)[0] + 1]) / 2
    return np.mean(np.abs((array2 - array1) / array2)) * 100

def MAPE(array1, array2):
    print('MAPE')
    print(datetime.now())
    for array in [array1, array2]:
        if not np.where(array == 0)[0].shape == (0,):
            print(array, 'Array contains 0 Value at line', np.where(array == 0)[0])
            for i in np.where(array == 0)[0]:
                array[i] = 0.0000025

    return np.mean(np.abs((array1 - array2) / array1)) * 100

def SMAPE(array1, array2):
    print('sMAPE')
    print(datetime.now())
    for array in [array1, array2]:
        if np.where(array == 0)[0].shape == (1,):
            print(array, 'Array contains 0 Value at line', np.where(array == 0)[0])
            array[np.where(array == 0)[0]] = (array[np.where(array == 0)[0] - 1] + array[np.where(array == 0)[0] + 1]) / 2
    return np.mean(np.abs(array1 - array2) / ((np.abs(array1)+np.abs(array2))/2)) * 100

def MSE(array1, array2):
    print('MSE')
    print(datetime.now())
    return np.mean((array1 - array2)**2)

def RSSE(array1, array2):
    print('RSSE')
    print(datetime.now())
    return (np.sum((array1 - array2)**2))**(1/2)


def Frechet(array1, array2):
    print('Frechet')
    print(datetime.now())
    df1 = np.zeros((wl_int,2))
    df2 = np.zeros((wl_int,2))
    base = np.arange(wl_int[0], wl_int[1]+5, 5)
    df1[:,0] = base
    df2[:,0] = base
    df1[:,0] = array1
    df2[:,0] = array2
    return similaritymeasures.frechet_dist(array1, array2)


def nan_half(array):
    ii=1
    for jj in np.arange(array.shape[0]):
        array[jj,0:ii]= np.nan
        ii+=1
    return array


def add_worksheet(workbook, name, op_names, integral_1, mean, sd, main_array1):
    worksheet1 = workbook.add_worksheet(name)
    row = 0
    col = 0
    row_2 = 0
    col_2 = 0

    for idx,i in enumerate(op_names):
        worksheet1.write(row + 1, col+  6, i)
        worksheet1.write(row + 2, col + 6, integral_1[idx])
        col += 1


    mad = compare_pairwise(main_array1, 'MAD')
    rmse = compare_pairwise(main_array1, 'RMSE')
    #frechet = compare_pairwise(main_array1, 'Frechet')
    mse = compare_pairwise(main_array1, 'MSE')
    rsse = compare_pairwise(main_array1, 'RSSE')
    mape= compare_pairwise(main_array1, 'MAPE')
    smape = compare_pairwise(main_array1, 'SMAPE')
    rmape = compare_pairwise(main_array1, 'rMAPE')

    for idx,i in enumerate(mad):
        for idx2, ii in enumerate(i):

            if not ii == 0:
                #worksheet1.write(10, col_2 + 6, frechet[idx,idx2], cell_format1)
                worksheet1.write(13, col_2 + 6, mad[idx, idx2], cell_format1)
                worksheet1.write(16, col_2 + 6, rmape[idx, idx2], cell_format1)
                worksheet1.write(19, col_2 + 6, smape[idx, idx2], cell_format1)
                worksheet1.write(22, col_2 + 6, mape[idx, idx2], cell_format1)
                worksheet1.write(25, col_2 + 6, mse[idx, idx2], cell_format1)
                worksheet1.write(28, col_2 + 6, rmse[idx, idx2], cell_format1)
                worksheet1.write(31, col_2 + 6, rsse[idx, idx2], cell_format1)
            col_2 += 1

    worksheet1.write(2, 4, mean, cell_format1)
    worksheet1.write(3, 4, sd, cell_format1)
    worksheet1.write(1, 1, 'Intergral & AUC', cell_format)
    worksheet1.write(2, 2, 'Integral', cell_format1)
    worksheet1.write(2, 3, 'mean', cell_format1)
    worksheet1.write(3, 3, 'SD', cell_format1)
    worksheet1.write(5, 2, 'Area under the curve', cell_format1)
    worksheet1.write(5, 3, 'mean', cell_format1)
    worksheet1.write(6, 3, 'SD', cell_format1)
    worksheet1.write(9, 1, 'heterogneity', cell_format)
    worksheet1.write(10, 2, 'Frechet', cell_format1)
    #worksheet1.write(10, 4, np,nanmean(frechet), cell_format1)
    #worksheet1.write(11, 4, np.nanstd(frechet), cell_format1)
    worksheet1.write(10, 3, 'mean', cell_format1)
    worksheet1.write(11, 3, 'SD', cell_format1)
    worksheet1.write(13, 2, 'MAD', cell_format1)
    worksheet1.write(13, 4, np.nanmean(mad), cell_format1)
    worksheet1.write(14, 4, np.nanstd(mad), cell_format1)
    worksheet1.write(13, 3, 'mean', cell_format1)
    worksheet1.write(14, 3, 'SD', cell_format1)
    worksheet1.write(16, 2, 'MAPE reverse', cell_format1)
    worksheet1.write(16, 4, np.nanmean(rmape), cell_format1)
    worksheet1.write(17, 4, np.nanstd(rmape), cell_format1)
    worksheet1.write(16, 3, 'mean', cell_format1)
    worksheet1.write(17, 3, 'SD', cell_format1)
    worksheet1.write(19, 2, 'MAPE symmetric', cell_format1)
    worksheet1.write(19, 4, np.nanmean(smape), cell_format1)
    worksheet1.write(20, 4, np.nanstd(smape), cell_format1)
    worksheet1.write(19, 3, 'mean', cell_format1)
    worksheet1.write(20, 3, 'SD', cell_format1)
    worksheet1.write(22, 2, 'MAPE', cell_format1)
    worksheet1.write(22, 4, np.nanmean(mape), cell_format1)
    worksheet1.write(23, 4, np.nanstd(mape), cell_format1)
    worksheet1.write(22, 3, 'mean', cell_format1)
    worksheet1.write(23, 3, 'SD', cell_format1)
    worksheet1.write(25, 2, 'MSE', cell_format1)
    worksheet1.write(25, 4, np.nanmean(mse), cell_format1)
    worksheet1.write(26, 4, np.nanstd(mse), cell_format1)
    worksheet1.write(25, 3, 'mean', cell_format1)
    worksheet1.write(26, 3, 'SD', cell_format1)
    worksheet1.write(28, 2, 'RMSE', cell_format1)
    worksheet1.write(28, 4, np.nanmean(rmse), cell_format1)
    worksheet1.write(29, 4, np.nanstd(rmse), cell_format1)
    worksheet1.write(28, 3, 'mean', cell_format1)
    worksheet1.write(29, 3, 'SD', cell_format1)
    worksheet1.write(31, 2, 'RSSE', cell_format1)
    worksheet1.write(31, 4, np.nanmean(rsse), cell_format1)
    worksheet1.write(32, 4, np.nanstd(rsse), cell_format1)
    worksheet1.write(31, 3, 'mean', cell_format1)
    worksheet1.write(32, 3, 'SD', cell_format1)
    worksheet1.set_column(0, 0, 3)
    worksheet1.set_column(1, 1, 10)
    worksheet1.set_column(2, 2, 15)
    worksheet1.freeze_panes(0, 4)
    worksheet1.write(2, 1, '=COUNTA(F3:AAA3)')

    return workbook


add_worksheet(workbook, 'Reflectance_0_derivative', integrate(get_array(file, "ref", 0, False))[0], integrate(get_array(file, "ref", 0, False))[1],
                integrate(get_array(file, "ref", 0, False))[2], integrate(get_array(file, "ref", 0, False))[3], integrate(get_array(file, "ref", 0, False))[4])
add_worksheet(workbook, 'Reflectance_0_dv_1_norm', integrate(get_array(file, "ref", 0, True))[0], integrate(get_array(file, "ref", 0, True))[1],
                integrate(get_array(file, "ref", 0, True))[2], integrate(get_array(file, "ref", 0, True))[3], integrate(get_array(file, "ref", 0, True))[4])

add_worksheet(workbook, 'Reflectance_1_derivative', integrate(get_array(file, "ref", 1, False))[0], integrate(get_array(file, "ref", 1, False))[1],
                integrate(get_array(file, "ref", 1, False))[2], integrate(get_array(file, "ref", 1, False))[3], integrate(get_array(file, "ref", 1, False))[4])
add_worksheet(workbook, 'Reflectance_1_dv_1_norm', integrate(get_array(file, "ref", 1, True))[0], integrate(get_array(file, "ref", 1, True))[1],
                integrate(get_array(file, "ref", 1, True))[2], integrate(get_array(file, "ref", 1, True))[3], integrate(get_array(file, "ref", 1, True))[4])

add_worksheet(workbook, 'Reflectance_2_derivative', integrate(get_array(file, "ref", 2, False))[0], integrate(get_array(file, "ref", 2, False))[1],
                integrate(get_array(file, "ref", 2, False))[2], integrate(get_array(file, "ref", 2, False))[3], integrate(get_array(file, "ref", 2, False))[4])
add_worksheet(workbook, 'Reflectance_2_dv_1_norm', integrate(get_array(file, "ref", 2, True))[0], integrate(get_array(file, "ref", 2, True))[1],
                integrate(get_array(file, "ref", 2, True))[2], integrate(get_array(file, "ref", 2, True))[3], integrate(get_array(file, "ref", 2, True))[4])

add_worksheet(workbook, 'Absorbance_0_derivative', integrate(get_array(file, "ab", 0, False))[0], integrate(get_array(file, "ab", 0, False))[1],
                integrate(get_array(file, "ab", 0, False))[2], integrate(get_array(file, "ab", 0, False))[3], integrate(get_array(file, "ab", 0, False))[4])
add_worksheet(workbook, 'Absorbance_0_dv_l1_norm', integrate(get_array(file, "ab", 0, True))[0], integrate(get_array(file, "ab", 0, True))[1],
                integrate(get_array(file, "ab", 0, True))[2], integrate(get_array(file, "ab", 0, True))[3], integrate(get_array(file, "ab", 0, True))[4])

add_worksheet(workbook, 'Absorbance_1_derivative', integrate(get_array(file, "ab", 1, False))[0], integrate(get_array(file, "ab", 1, False))[1],
                integrate(get_array(file, "ab", 1, False))[2], integrate(get_array(file, "ab", 1, False))[3], integrate(get_array(file, "ab", 1, False))[4])
add_worksheet(workbook, 'Absorbance_1_dv_l1_norm', integrate(get_array(file, "ab", 1, True))[0], integrate(get_array(file, "ab", 1, True))[1],
                integrate(get_array(file, "ab",1, True))[2], integrate(get_array(file, "ab", 1, True))[3], integrate(get_array(file, "ab", 1, True))[4])

add_worksheet(workbook, 'Absorbance_2_derivative', integrate(get_array(file, "ab", 2, False))[0], integrate(get_array(file, "ab", 2, False))[1],
                integrate(get_array(file, "ab", 2, False))[2], integrate(get_array(file, "ab", 2, False))[3], integrate(get_array(file, "ab", 2, False))[4])
add_worksheet(workbook, 'Absorbance_2_dv_l1_norm', integrate(get_array(file, "ab", 2, True))[0], integrate(get_array(file, "ab", 2, True))[1],
                integrate(get_array(file, "ab", 2, True))[2], integrate(get_array(file, "ab", 2, True))[3], integrate(get_array(file, "ab", 2, True))[4])

workbook.close()
