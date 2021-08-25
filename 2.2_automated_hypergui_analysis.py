"""
@author: Alexander Studier-Fischer, Jan Odenthal, Berkin Özdemir, University of Heidelberg
"""
# -*- coding: utf-8 -*-

mkSto2 =True
mkNir =True
mkThi =True
mkTwi =True
mkOhi =True
mkTli =True

import numpy as np
import matplotlib.pyplot as plt
import skimage.color
import os
import logging
import xlsxwriter
import math
import matplotlib.pyplot as plt
import matplotlib
import glob
import imageio
from PIL import Image


file = [
["103","0","0","100"],
["105","0","0","100"],
["106","0","0","100"],
["108","0","0","99"],
["110","0","0","99"],
["111","0","0","99"],
["113","0","0","99"],
["115","0","0","98"],
["116","0","0","98"],
['118', '0', '0', '98'],
['119', '0', '0', '98'],
['121', '0', '0', '98'],
['123', '0', '0', '97'],
['124', '0', '0', '97'],
['127', '0', '0', '97'],
['129', '0', '0', '97'],
['130', '0', '0', '97'],
['132', '0', '0', '96'],
['134', '0', '0', '96'],
['135', '0', '0', '96'],
['137', '0', '0', '96'],
['138', '0', '0', '96'],
['140', '0', '0', '95'],
['142', '0', '0', '95'],
['143', '0', '0', '95'],
['145', '0', '0', '95'],
['147', '0', '0', '94'],
['148', '0', '0', '94'],
['150', '0', '0', '94'],
['152', '0', '0', '94'],
['153', '0', '0', '94'],
['155', '0', '0', '93'],
['156', '0', '0', '93'],
['158', '0', '0', '93'],
['160', '0', '0', '93'],
['161', '0', '0', '93'],
['163', '0', '0', '92'],
['165', '0', '0', '92'],
['166', '0', '0', '92'],
['168', '0', '0', '92'],
['170', '0', '0', '92'],
['171', '0', '0', '92'],
['173', '0', '0', '91'],
['174', '0', '0', '91'],
['176', '0', '0', '91'],
['178', '0', '0', '90'],
['179', '0', '0', '90'],
['181', '0', '0', '90'],
['183', '0', '0', '90'],
['184', '0', '0', '90'],
['186', '0', '0', '89'],
['188', '0', '0', '89'],
['189', '0', '0', '89'],
['191', '0', '0', '89'],
['192', '0', '0', '89'],
['194', '0', '0', '88'],
['196', '0', '0', '88'],
['197', '0', '0', '88'],
['199', '0', '0', '88'],
['201', '0', '0', '88'],
['202', '0', '0', '88'],
['204', '0', '0', '87'],
['206', '0', '0', '87'],
['207', '0', '0', '87'],
['209', '0', '0', '86'],
['210', '0', '0', '86'],
['212', '0', '0', '86'],
['214', '0', '0', '86'],
['215', '0', '0', '86'],
['217', '0', '0', '85'],
['220', '0', '0', '85'],
['221', '0', '0', '85'],
['223', '0', '0', '85'],
['224', '0', '0', '85'],
['226', '0', '0', '84'],
['228', '0', '0', '84'],
['229', '0', '0', '84'],
['231', '0', '0', '84'],
['233', '0', '0', '84'],
['234', '0', '0', '83'],
['236', '0', '0', '83'],
['238', '0', '0', '83'],
['239', '0', '0', '83'],
['241', '0', '0', '83'],
['242', '0', '0', '82'],
['244', '0', '0', '82'],
['246', '0', '0', '82'],
['247', '0', '0', '82'],
['249', '0', '0', '82'],
['251', '0', '0', '81'],
['252', '0', '0', '81'],
['254', '0', '0', '81'],
['255', '0', '0', '81'],
['255', '0', '0', '80'],
['255', '0', '0', '80'],
['255', '0', '0', '80'],
['255', '0', '0', '80'],
['255', '0', '0', '80'],
['255', '0', '0', '80'],
['255', '2', '0', '79'],
['255', '6', '0', '79'],
['255', '10', '0', '79'],
['255', '15', '0', '79'],
['255', '18', '0', '78'],
['255', '22', '0', '78'],
['255', '25', '0', '78'],
['255', '28', '0', '78'],
['255', '31', '0', '78'],
['255', '35', '0', '77'],
['255', '37', '0', '77'],
['255', '41', '0', '77'],
['255', '44', '0', '77'],
['255', '46', '0', '76'],
['255', '49', '0', '76'],
['255', '53', '0', '76'],
['255', '56', '0', '76'],
['255', '58', '0', '76'],
['255', '61', '0', '76'],
['255', '64', '0', '75'],
['255', '66', '0', '75'],
['255', '70', '0', '75'],
['255', '73', '0', '74'],
['255', '75', '0', '74'],
['255', '78', '0', '74'],
['255', '81', '0', '74'],
['255', '83', '0', '74'],
['255', '86', '0', '73'],
['255', '89', '0', '73'],
['255', '91', '0', '73'],
['255', '94', '0', '73'],
['255', '98', '0', '73'],
['255', '101', '0', '72'],
['255', '103', '0', '72'],
['255', '106', '0', '72'],
['255', '109', '0', '72'],
['255', '111', '0', '72'],
['255', '114', '0', '72'],
['255', '117', '0', '71'],
['255', '119', '0', '71'],
['255', '122', '0', '71'],
['255', '125', '0', '70'],
['255', '127', '0', '70'],
['255', '130', '0', '70'],
['255', '133', '0', '70'],
['255', '136', '0', '70'],
['255', '138', '0', '69'],
['255', '141', '0', '69'],
['255', '145', '0', '69'],
['255', '147', '0', '69'],
['255', '150', '0', '69'],
['255', '153', '0', '68'],
['255', '155', '0', '68'],
['255', '158', '0', '68'],
['255', '161', '0', '68'],
['255', '163', '0', '68'],
['255', '166', '0', '67'],
['255', '169', '0', '67'],
['255', '171', '0', '67'],
['255', '174', '0', '67'],
['255', '177', '0', '66'],
['255', '180', '0', '66'],
['255', '182', '0', '66'],
['255', '185', '0', '66'],
['255', '188', '0', '66'],
['255', '190', '0', '66'],
['255', '193', '0', '65'],
['255', '196', '0', '65'],
['255', '198', '0', '65'],
['255', '201', '0', '65'],
['255', '204', '0', '64'],
['255', '206', '0', '64'],
['255', '210', '0', '64'],
['255', '213', '0', '64'],
['255', '216', '0', '64'],
['255', '218', '0', '63'],
['255', '221', '0', '63'],
['255', '224', '0', '63'],
['255', '226', '0', '63'],
['255', '229', '0', '62'],
['255', '232', '0', '62'],
['255', '234', '0', '62'],
['255', '237', '0', '62'],
['255', '240', '0', '62'],
['255', '242', '0', '61'],
['255', '245', '0', '61'],
['255', '248', '0', '61'],
['255', '251', '0', '61'],
['255', '253', '0', '61'],
['255', '255', '0', '60'],
['253', '255', '0', '60'],
['251', '255', '0', '60'],
['248', '255', '0', '60'],
['245', '255', '0', '60'],
['243', '255', '0', '59'],
['240', '255', '0', '59'],
['237', '255', '0', '59'],
['235', '255', '0', '59'],
['232', '255', '0', '58'],
['229', '255', '0', '58'],
['227', '255', '0', '58'],
['223', '255', '0', '58'],
['220', '255', '0', '58'],
['217', '255', '0', '58'],
['215', '255', '0', '57'],
['212', '255', '0', '57'],
['209', '255', '0', '57'],
['207', '255', '0', '57'],
['204', '255', '0', '56'],
['201', '255', '0', '56'],
['199', '255', '0', '56'],
['196', '255', '0', '56'],
['192', '255', '0', '56'],
['190', '255', '0', '56'],
['187', '255', '0', '55'],
['184', '255', '0', '55'],
['181', '255', '0', '55'],
['179', '255', '0', '54'],
['176', '255', '0', '54'],
['173', '255', '0', '54'],
['171', '255', '0', '54'],
['168', '255', '0', '54'],
['164', '255', '0', '54'],
['162', '255', '0', '53'],
['159', '255', '0', '53'],
['156', '255', '0', '53'],
['154', '255', '0', '53'],
['151', '255', '0', '52'],
['148', '255', '0', '52'],
['145', '255', '0', '52'],
['142', '255', '0', '52'],
['139', '255', '0', '52'],
['136', '255', '0', '51'],
['134', '255', '0', '51'],
['131', '255', '0', '51'],
['127', '255', '0', '51'],
['125', '255', '0', '50'],
['122', '255', '0', '50'],
['119', '255', '0', '50'],
['116', '255', '0', '50'],
['113', '255', '0', '50'],
['110', '255', '0', '50'],
['108', '255', '0', '49'],
['104', '255', '0', '49'],
['101', '255', '0', '49'],
['98', '255', '0', '48'],
['95', '255', '0', '48'],
['92', '255', '0', '48'],
['89', '255', '0', '48'],
['86', '255', '0', '48'],
['83', '255', '0', '48'],
['79', '255', '0', '47'],
['77', '255', '0', '47'],
['73', '255', '0', '47'],
['70', '255', '0', '47'],
['67', '255', '0', '46'],
['64', '255', '0', '46'],
['60', '255', '0', '46'],
['57', '255', '0', '46'],
['53', '255', '0', '46'],
['49', '255', '0', '46'],
['45', '255', '0', '45'],
['42', '255', '0', '45'],
['37', '255', '0', '45'],
['32', '255', '0', '45'],
['28', '255', '0', '44'],
['22', '255', '0', '44'],
['14', '255', '0', '44'],
['8', '255', '0', '44'],
['0', '255', '0', '44'],
['0', '255', '0', '44'],
['0', '255', '0', '43'],
['0', '255', '0', '43'],
['0', '255', '0', '43'],
['0', '255', '0', '42'],
['0', '255', '0', '42'],
['0', '255', '0', '42'],
['0', '255', '0', '42'],
['0', '255', '0', '42'],
['0', '255', '0', '42'],
['0', '255', '0', '41'],
['0', '255', '0', '41'],
['0', '255', '0', '41'],
['0', '255', '0', '40'],
['0', '255', '0', '40'],
['0', '255', '0', '40'],
['0', '252', '0', '40'],
['0', '250', '0', '40'],
['0', '247', '0', '39'],
['0', '244', '0', '39'],
['0', '241', '0', '39'],
['0', '239', '0', '39'],
['0', '236', '0', '39'],
['0', '233', '0', '38'],
['0', '231', '0', '38'],
['0', '228', '0', '38'],
['0', '225', '0', '38'],
['0', '223', '0', '38'],
['0', '220', '0', '37'],
['0', '217', '4', '37'],
['0', '215', '12', '37'],
['0', '212', '21', '37'],
['0', '209', '27', '36'],
['0', '206', '33', '36'],
['0', '204', '37', '36'],
['0', '201', '42', '36'],
['0', '198', '46', '36'],
['0', '196', '49', '36'],
['0', '193', '54', '35'],
['0', '190', '58', '35'],
['0', '188', '60', '35'],
['0', '185', '64', '35'],
['0', '182', '68', '34'],
['0', '180', '71', '34'],
['0', '177', '74', '34'],
['0', '173', '78', '34'],
['0', '170', '81', '34'],
['0', '168', '84', '34'],
['0', '165', '87', '33'],
['0', '162', '91', '33'],
['0', '160', '93', '33'],
['0', '157', '96', '32'],
['0', '154', '99', '32'],
['0', '152', '102', '32'],
['0', '149', '105', '32'],
['0', '146', '108', '32'],
['0', '144', '110', '32'],
['0', '141', '114', '31'],
['0', '138', '117', '31'],
['0', '136', '119', '31'],
['0', '132', '122', '31'],
['0', '129', '126', '30'],
['0', '126', '129', '30'],
['0', '124', '131', '30'],
['0', '121', '134', '30'],
['0', '118', '137', '30'],
['0', '116', '139', '29'],
['0', '113', '142', '29'],
['0', '109', '146', '29'],
['0', '107', '148', '29'],
['0', '104', '151', '28'],
['0', '101', '154', '28'],
['0', '99', '156', '28'],
['0', '96', '159', '28'],
['0', '92', '162', '28'],
['0', '89', '165', '28'],
['0', '87', '167', '28'],
['0', '84', '171', '27'],
['0', '80', '174', '27'],
['0', '78', '176', '27'],
['0', '75', '179', '27'],
['0', '71', '182', '26'],
['0', '69', '184', '26'],
['0', '66', '187', '26'],
['0', '62', '190', '26'],
['0', '60', '192', '25'],
['0', '56', '195', '25'],
['0', '52', '198', '25'],
['0', '50', '200', '24'],
['0', '46', '203', '24'],
['0', '42', '207', '24'],
['0', '37', '210', '24'],
['0', '34', '212', '24'],
['0', '29', '215', '24'],
['0', '24', '218', '24'],
['0', '20', '220', '23'],
['0', '13', '223', '23'],
['0', '4', '226', '23'],
['0', '0', '228', '23'],
['0', '0', '231', '22'],
['0', '0', '234', '22'],
['0', '0', '236', '22'],
['0', '0', '239', '22'],
['0', '0', '242', '22'],
['0', '0', '245', '21'],
['0', '0', '247', '21'],
['0', '0', '250', '21'],
['0', '0', '253', '21'],
['0', '0', '255', '20'],
['0', '0', '255', '20'],
['0', '0', '255', '20'],
['0', '0', '255', '20'],
['0', '0', '254', '20'],
['0', '0', '252', '19'],
['0', '0', '251', '19'],
['0', '0', '249', '19'],
['0', '0', '247', '19'],
['0', '0', '246', '19'],
['0', '0', '244', '18'],
['0', '0', '242', '18'],
['0', '0', '241', '18'],
['0', '0', '239', '18'],
['0', '0', '238', '18'],
['0', '0', '236', '18'],
['0', '0', '234', '17'],
['0', '0', '233', '17'],
['0', '0', '231', '17'],
['0', '0', '229', '16'],
['0', '0', '228', '16'],
['0', '0', '226', '16'],
['0', '0', '224', '16'],
['0', '0', '223', '16'],
['0', '0', '221', '16'],
['0', '0', '220', '15'],
['0', '0', '218', '15'],
['0', '0', '216', '15'],
['0', '0', '215', '15'],
['0', '0', '213', '14'],
['0', '0', '211', '14'],
['0', '0', '210', '14'],
['0', '0', '208', '14'],
['0', '0', '207', '14'],
['0', '0', '205', '13'],
['0', '0', '203', '13'],
['0', '0', '202', '13'],
['0', '0', '200', '13'],
['0', '0', '198', '12'],
['0', '0', '196', '12'],
['0', '0', '194', '12'],
['0', '0', '192', '12'],
['0', '0', '191', '12'],
['0', '0', '189', '12'],
['0', '0', '188', '11'],
['0', '0', '186', '11'],
['0', '0', '184', '11'],
['0', '0', '183', '11'],
['0', '0', '181', '10'],
['0', '0', '179', '10'],
['0', '0', '178', '10'],
['0', '0', '176', '10'],
['0', '0', '174', '10'],
['0', '0', '173', '9'],
['0', '0', '171', '9'],
['0', '0', '170', '9'],
['0', '0', '168', '9'],
['0', '0', '166', '8'],
['0', '0', '165', '8'],
['0', '0', '163', '8'],
['0', '0', '161', '8'],
['0', '0', '160', '8'],
['0', '0', '158', '8'],
['0', '0', '156', '7'],
['0', '0', '155', '7'],
['0', '0', '153', '7'],
['0', '0', '152', '7'],
['0', '0', '150', '6'],
['0', '0', '148', '6'],
['0', '0', '147', '6'],
['0', '0', '145', '6'],
['0', '0', '143', '6'],
['0', '0', '142', '5'],
['0', '0', '140', '5'],
['0', '0', '138', '5'],
['0', '0', '136', '5'],
['0', '0', '134', '4'],
['0', '0', '133', '4'],
['0', '0', '131', '4'],
['0', '0', '129', '4'],
['0', '0', '128', '4'],
['0', '0', '126', '3'],
['0', '0', '124', '3'],
['0', '0', '123', '3'],
['0', '0', '121', '3'],
['0', '0', '119', '3'],
['0', '0', '118', '2'],
['0', '0', '116', '2'],
['0', '0', '115', '2'],
['0', '0', '113', '2'],
['0', '0', '111', '2'],
['0', '0', '110', '2'],
['0', '0', '108', '1'],
['0', '0', '106', '1'],
['0', '0', '105', '1'],
['0', '0', '103', '1'],
['0', '0', '103', '1'],
['0', '0', '0', '0']]

Nero = False
spec_tup1 = (False, True, False)
spec_tup5 = (True, True, False)
base = os.getcwd()
data_path = base

NDIM = 3
rgb_list = []
scale_list = []
for nums in file:
    rgb_list.append((int(nums[0]), int(nums[1]), int(nums[2])))
    scale_list.append(int(nums[3]))

a = np.asarray(rgb_list)
a.shape = int(a.size / NDIM), NDIM

RGB_FILE = "_RGB-Image.png"
STO2_FILE = "_Oxygenation.png"
NIR_FILE = "_NIR-Perfusion.png"
THI_FILE = "_THI.png"
TWI_FILE = "_TWI.png"
TLI_FILE = "_TLI.png"
OHI_FILE = "_OHI.png"

def image_to_array(filename):
    return imageio.imread(filename)
        
def find_closest_3d(point):
    d = ((a - point) ** 2).sum(axis=1)
    ndx = d.argsort()
    return scale_list[ndx[0]]
        
def rgb_image_to_hsi_array(img_array):
    array = []
    truth = isinstance(img_array, np.ma.MaskedArray)
    if truth:
        mask = img_array.mask[:, :, 0]
    # iterate over pixels in image
    for i in range(len(img_array)):
        for j in range(len(img_array[i])):
            # normalise rgb values
            zero = img_array[i][j][0]
            one = img_array[i][j][1]
            two = img_array[i][j][2]
            if truth:
                if not mask[i][j]:
                    closest = find_closest_3d((zero, one, two))
                    array.append(closest)
                else:
                    array.append(str('NaN'))
            else:
                array.append(find_closest_3d((zero, one, two)))
    return np.asarray(array).reshape((480, 640))

def get_channels(mask, histogram_data):
    if mkTwi:
        TWI_data = histogram_data.get_twi_og()
        twi = calc_data(mask, TWI_data)
    else:
        twi = None
    if mkSto2:
        STO2_data = histogram_data.get_sto2_og()
        sto2 = calc_data(mask, STO2_data)
    else:
        sto2 = None
    if mkThi:
        THI_data = histogram_data.get_thi_og()
        thi = calc_data(mask, THI_data)
    else:
        thi = None
    if mkNir:
        NIR_data = histogram_data.get_nir_og()
        nir = calc_data(mask, NIR_data)
    else:
        nir = None
    if mkOhi:
        OHI_data = histogram_data.get_ohi_og()
        ohi = calc_data(mask, OHI_data)
    else:
        ohi = None
    if mkTli:
        TLI_data = histogram_data.get_tli_og()
        tli = calc_data(mask, TLI_data)
    else:
        tli = None
        
    return {"twi": twi, "thi": thi, "nir": nir, "sto2": sto2, "tli": tli, "ohi": ohi}

def get_current_original_data(mask, data):
    data = np.asarray(rgb_image_to_hsi_array(data))
    data = np.ma.array(data, mask=np.rot90(mask))
    return data

def calc_data(mask, data):
    if data is not None:
        stats_data = get_current_original_data(mask, data).flatten()
        data = np.ma.sort(stats_data)
        length = np.ma.count(data)
        logging.debug("CALCULATING STATS...")
        mean_value = np.round(np.ma.mean(data), 4)
        sd_value = np.round(np.ma.std(data), 4)
        median_value = np.round(data[int(length*1/2)], 4)
        iqr_value = (np.round(data[int(length*1/4)], 4), round(data[int(length*3/4)], 4))
        iqr_low = iqr_value[0]
        iqr_high = iqr_value[1]
        min_value = np.round(np.ma.min(data), 4)
        max_value = np.round(np.ma.max(data), 4)
        return [mean_value, sd_value, median_value, iqr_low, iqr_high, min_value, max_value]
    else:
        return None

class AbsSpecAnalysis:
    # performs analyses necessary for the absorption spectrum
    def __init__(self, data_cube, wavelength, spec_tup, mask=None):


        # inputs
        self.data_cube = data_cube
        self.mask = mask
        self.wavelength = wavelength
        self.absorbance = bool(spec_tup[0])
        self.normal = not bool(spec_tup[1])
        self.negative = bool(spec_tup[2])

        # calculated generally
        self.x1 = None
        self.x2 = None
        self.x_absorbance = None
        self.x_reflectance = None
        self.x_absorbance_w = None
        self.x_reflectance_w = None
        self.x_absorbance_masked = None
        self.x_absorbance_masked_w = None
        self.x_reflectance_masked = None
        self.x_reflectance_masked_w = None

        # specific to module
        self.absorption_roi = None
        self.absorption_roi_masked = None

        # data cube 
        self.key = None
        self.value = None

        self.analysis()

    def analysis(self):
        self._calc_general()
        self._calc_absorption_spec()

    # --------------------------------------------------- UPDATERS ----------------------------------------------------

    def update_mask(self, new_mask):
        self.mask = new_mask
        self.analysis()

    def update_wavelength(self, new_wavelength):
        self.wavelength = new_wavelength
        self.analysis()

    def update_normal(self, new_normal):
        self.normal = new_normal
        self.analysis()

    def update_absorbance(self, new_absorbance):
        self.absorbance = new_absorbance
        self.analysis()

    # ------------------------------------------------- CALCULATORS --------------------------------------------------

    def _calc_absorption_spec(self):
        if self.absorbance:
            self.absorption_roi = self._calc_absorption_spec_roi(self.x_absorbance)
            if self.mask is not None:
                self.absorption_roi_masked = self._calc_absorption_spec_roi(self.x_absorbance_masked)
        else:
            self.absorption_roi = self._calc_absorption_spec_roi(self.x_reflectance)
            if self.mask is not None:
                self.absorption_roi_masked = self._calc_absorption_spec_roi(self.x_reflectance_masked)

    @staticmethod
    def _calc_absorption_spec_roi(data):
        absorption_roi = []
        wavelengths = np.arange(500, 1000, 5)

        for i in range(data.shape[2]):
            tmp = np.ma.median(data[:, :, i])
            absorption_roi.append((int(wavelengths[i]), tmp))

        return np.array(absorption_roi)

    # --------------------------------------------- GENERAL CALCULATORS ----------------------------------------------

    def _calc_general(self):
        self.__calc_x1()
        self.__calc_x_reflectance()
        self.__calc_x2()
        self.__calc_x_absorbance()

    def __calc_x1(self):
        neg = 0
        # normalise
        if self.normal and not self.absorbance:
            data = self.data_cube
            if np.ma.min(self.data_cube) < 0:
                neg = np.abs(np.ma.min(data))
                data = data + np.abs(np.ma.min(data))
            if np.ma.min(self.data_cube) > 0:
                data = data - np.abs(np.ma.min(data))
            neg = neg / np.ma.max(data)
            self.x1 = data / np.ma.max(data)
        else:
            self.x1 = self.data_cube
        # mask negatives
        if self.negative:
            self.x1 = np.ma.array(self.x1, mask=self.x1 < neg)

    def __calc_x_reflectance(self):
        self.x_reflectance = self.x1

        if self.wavelength[0] != self.wavelength[1]:
            wav_lower = int(round(max(0, min(self.wavelength)), 0))
            wav_upper = int(round(min(max(self.wavelength), 99), 0))
            self.x_reflectance_w = np.mean(self.x_reflectance[:, :, wav_lower:wav_upper+1], axis=2)
        else:
            self.x_reflectance_w = self.x_reflectance[:, :, self.wavelength[0]]

        if self.mask is not None:
            mask = np.array([self.mask.T] * 100).T
            self.x_reflectance_masked = np.ma.array(self.x_reflectance[:, :, :], mask=mask)
            # self.x_reflectance_masked_w = np.ma.array(self.x_reflectance[:, :, self.wavelength[0]], mask=self.mask)
            if self.wavelength[0] != self.wavelength[1]:
                wav_lower = int(round(max(0, min(self.wavelength)), 0))
                wav_upper = int(round(min(max(self.wavelength), 99), 0))
                self.x_reflectance_masked_w = np.ma.array(np.mean(self.x_reflectance[:, :, wav_lower:wav_upper+1],
                                                                  axis=2), mask=self.mask)
            else:
                self.x_reflectance_masked_w = np.ma.array(self.x_reflectance[:, :, self.wavelength[0]], mask=self.mask)

    def __calc_x2(self):
        self.x2 = -np.ma.log(self.x1)
        self.x2 = np.ma.array(self.x2, mask=~np.isfinite(self.x2))
        neg = 0
        # normalise
        if self.normal and self.absorbance:
            data = self.x2
            if np.ma.min(self.x2) < 0:
                neg = np.abs(np.ma.min(data))
                data = data + np.abs(np.ma.min(data))
            if np.ma.min(self.x2) > 0:
                data = data - np.abs(np.ma.min(data))
            neg = neg / np.ma.max(data)
            self.x2 = data / np.ma.max(data)
        # mask negatives
        if self.negative:
            self.x2 = np.ma.array(self.x2, mask=self.x2 < neg)

    def __calc_x_absorbance(self):
        self.x_absorbance = self.x2

        if self.wavelength[0] != self.wavelength[1]:
            wav_lower = int(round(max(0, min(self.wavelength)), 0))
            wav_upper = int(round(min(max(self.wavelength), 99), 0))
            self.x_absorbance_w = np.mean(self.x_absorbance[:, :, wav_lower:wav_upper+1], axis=2)
        else:
            self.x_absorbance_w = self.x_absorbance[:, :, self.wavelength[0]]

        if self.mask is not None:
            # self.x_absorbance_masked = self.__apply_2DMask_on_3DArray(self.mask, self.x_absorbance)
            mask = np.array([self.mask.T] * 100).T
            self.x_absorbance_masked = np.ma.array(self.x_absorbance[:, :, :], mask=mask)
            # self.x_absorbance_masked = np.ma.array(self.x_absorbance[:, :, :], mask=np.array([self.mask] * 100))
            if self.wavelength[0] != self.wavelength[1]:
                wav_lower = int(round(min(0, min(self.wavelength)), 0))
                wav_upper = int(round(max(max(self.wavelength), 99), 0))
                self.x_absorbance_masked_w = np.ma.array(np.mean(self.x_absorbance[:, :, wav_lower:wav_upper+1],
                                                                 axis=2), mask=self.mask)
            else:
                self.x_absorbance_masked_w = np.ma.array(self.x_absorbance[:, :, self.wavelength[0]], mask=self.mask)     
                

class HistogramAnalysis:
    def __init__(self, path, data_cube, wavelength, specs, mask=None):

        # inputs
        self.path = path
        self.data_cube = data_cube
        self.mask = mask
        self.wavelength = wavelength
        self.absorbance = bool(specs[0])
        self.normal = not bool(specs[1])
        self.negative = bool(specs[2])

        # calculated generally
        self.x1 = None
        self.x2 = None
        self.x_absorbance = None
        self.x_reflectance = None
        self.x_absorbance_w = None
        self.x_reflectance_w = None
        self.x_absorbance_masked = None
        self.x_absorbance_masked_w = None
        self.x_reflectance_masked = None
        self.x_reflectance_masked_w = None

        # specific to module
        self.rgb_og = None
        self.sto2_og = None
        self.nir_og = None
        self.thi_og = None
        self.twi_og = None
        self.histogram_data = None
        self.histogram_data_masked = None

        self.analysis()

    def analysis(self):
        self._calc_general()
        self._calc_histogram_data()

    # --------------------------------------------------- UPDATERS ----------------------------------------------------

    def update_mask(self, new_mask):
        self.mask = new_mask
        self.analysis()

    def update_wavelength(self, new_wavelength):
        self.wavelength = new_wavelength
        self.analysis()

    def update_normal(self, new_normal):
        self.normal = new_normal
        self.analysis()

    def update_absorbance(self, new_absorbance):
        self.absorbance = new_absorbance
        self.analysis()

    # --------------------------------------------------- GETTERS ----------------------------------------------------

    def ensure_shape(self, data):
        arr = np.asarray(data.shape)
        g = arr.copy()
        arr.sort()
        i1 = np.where(g == arr[1])[0][0]
        i2 = np.where(g == arr[2])[0][0]
        i3 = np.where(g == arr[0])[0][0]
        return np.moveaxis(data, [i1, i2, i3], [0, 1, 2])

    # ------------------------------------------------- CALCULATORS --------------------------------------------------

    def _calc_histogram_data(self):
        if self.absorbance:
            self.histogram_data = self.x2
            if self.mask is not None:
                self.histogram_data_masked = self.x_absorbance_masked
        else:
            self.histogram_data = self.x1
            if self.mask is not None:
                self.histogram_data_masked = self.x_reflectance_masked

    # --------------------------------------------- GENERAL CALCULATORS ----------------------------------------------

    def _calc_general(self):
        self.__calc_x1()
        self.__calc_x_reflectance()
        self.__calc_x2()
        self.__calc_x_absorbance()

    def __calc_x1(self):
        neg = 0
        # normalise
        if self.normal and not self.absorbance:
            data = self.data_cube
            if np.ma.min(self.data_cube) < 0:
                neg = np.abs(np.ma.min(data))
                data = data + np.abs(np.ma.min(data))
            if np.ma.min(self.data_cube) > 0:
                data = data - np.abs(np.ma.min(data))
            neg = neg / np.ma.max(data)
            self.x1 = data / np.ma.max(data)
        else:
            self.x1 = self.data_cube
        # mask negatives
        if self.negative:
            self.x1 = np.ma.array(self.x1, mask=self.x1 < neg)

    def __calc_x_reflectance(self):
        self.x_reflectance = self.x1

        if self.wavelength[0] != self.wavelength[1]:
            wav_lower = int(round(max(0, min(self.wavelength)), 0))
            wav_upper = int(round(min(max(self.wavelength), 99), 0))
            self.x_reflectance_w = np.mean(self.x_reflectance[:, :, wav_lower:wav_upper+1], axis=2)
        else:
            self.x_reflectance_w = self.x_reflectance[:, :, self.wavelength[0]]

        if self.mask is not None:
            mask = np.array([self.mask.T] * 100).T
            self.x_reflectance_masked = np.ma.array(self.x_reflectance[:, :, :], mask=mask)
            # self.x_reflectance_masked_w = np.ma.array(self.x_reflectance[:, :, self.wavelength[0]], mask=self.mask)
            if self.wavelength[0] != self.wavelength[1]:
                wav_lower = int(round(max(0, min(self.wavelength)), 0))
                wav_upper = int(round(min(max(self.wavelength), 99), 0))
                self.x_reflectance_masked_w = np.ma.array(np.mean(self.x_reflectance[:, :, wav_lower:wav_upper+1],
                                                                  axis=2), mask=self.mask)
            else:
                self.x_reflectance_masked_w = np.ma.array(self.x_reflectance[:, :, self.wavelength[0]], mask=self.mask)

    def __calc_x2(self):
        self.x2 = -np.ma.log(self.x1)
        self.x2 = np.ma.array(self.x2, mask=~np.isfinite(self.x2))
        neg = 0
        # normalise
        if self.normal and self.absorbance:
            data = self.x2
            if np.ma.min(self.x2) < 0:
                neg = np.abs(np.ma.min(data))
                data = data + np.abs(np.ma.min(data))
            if np.ma.min(self.x2) > 0:
                data = data - np.abs(np.ma.min(data))
            neg = neg / np.ma.max(data)
            self.x2 = data / np.ma.max(data)
        # mask negatives
        if self.negative:
            self.x2 = np.ma.array(self.x2, mask=self.x2 < neg)

    def __calc_x_absorbance(self):
        self.x_absorbance = self.x2

        if self.wavelength[0] != self.wavelength[1]:
            wav_lower = int(round(max(0, min(self.wavelength)), 0))
            wav_upper = int(round(min(max(self.wavelength), 99), 0))
            self.x_absorbance_w = np.mean(self.x_absorbance[:, :, wav_lower:wav_upper+1], axis=2)
        else:
            self.x_absorbance_w = self.x_absorbance[:, :, self.wavelength[0]]

        if self.mask is not None:
            # self.x_absorbance_masked = self.__apply_2DMask_on_3DArray(self.mask, self.x_absorbance)
            mask = np.array([self.mask.T] * 100).T
            self.x_absorbance_masked = np.ma.array(self.x_absorbance[:, :, :], mask=mask)
            # self.x_absorbance_masked = np.ma.array(self.x_absorbance[:, :, :], mask=np.array([self.mask] * 100))
            if self.wavelength[0] != self.wavelength[1]:
                wav_lower = int(round(min(0, min(self.wavelength)), 0))
                wav_upper = int(round(max(max(self.wavelength), 99), 0))
                self.x_absorbance_masked_w = np.ma.array(np.mean(self.x_absorbance[:, :, wav_lower:wav_upper+1],
                                                                 axis=2), mask=self.mask)
            else:
                self.x_absorbance_masked_w = np.ma.array(self.x_absorbance[:, :, self.wavelength[0]], mask=self.mask)
                
    def get_sto2_og(self):
        filename = str(self.path[:-13]) + STO2_FILE
        if os.path.exists(filename) and mkSto2:
            self.sto2_og = image_to_array(filename)
            if self.sto2_og.shape[0] == 550:
                chopped = self.sto2_og[50:530, 50:690, :3]
            else:
                chopped = self.sto2_og[26:506, 4:644, :3]
        else:
            chopped = None
        return chopped

    def get_nir_og(self):
        filename = str(self.path[:-13]) + NIR_FILE
        if os.path.exists(filename) and mkNir:
            self.nir_og = image_to_array(filename)
            if self.nir_og.shape[0] == 550:
                chopped = self.nir_og[50:530, 50:690, :3]
            else:
                chopped = self.nir_og[26:506, 4:644, :3]
        else:
            chopped = None
        return chopped

    def get_thi_og(self):
        filename = str(self.path[:-13]) + THI_FILE
        if os.path.exists(filename) and mkThi:
            self.thi_og = image_to_array(filename)
            if self.thi_og.shape[0] == 550:
                chopped = self.thi_og[50:530, 50:690, :3]
            else:
                chopped = self.thi_og[26:506, 4:644, :3]
        else:
            chopped = None
        return chopped

    def get_twi_og(self):
        filename = str(self.path[:-13]) + TWI_FILE
        if os.path.exists(filename) and mkTwi:
            self.twi_og = image_to_array(filename)
            if self.twi_og.shape[0] == 550:
                chopped = self.twi_og[50:530, 50:690, :3]
            else:
                chopped = self.twi_og[26:506, 4:644, :3]
        else:
            chopped = None
        return chopped
    
    def get_tli_og(self):
        filename = str(self.path[:-13]) + TLI_FILE
        if os.path.exists(filename) and mkTli:
            self.tli_og = image_to_array(filename)
            if self.tli_og.shape[0] == 550:
                chopped = self.tli_og[50:530, 50:690, :3]
            else:
                chopped = self.tli_og[26:506, 4:644, :3]
        else:
            chopped = None
        return chopped

    def get_ohi_og(self):
        filename = str(self.path[:-13]) + OHI_FILE
        if os.path.exists(filename) and mkOhi:
            self.ohi_og = image_to_array(filename)
            if self.ohi_og.shape[0] == 550:
                chopped = self.ohi_og[50:530, 50:690, :3]
            else:
                chopped = self.ohi_og[26:506, 4:644, :3]
        else:
            chopped = None
        return chopped


                
                
                
def save_absorption_spec(output_path, AbsAna, spec_number):
    data = AbsAna.absorption_roi_masked[:, 1]
    save_absorption_spec_graph(output_path, data, spec_number, True, True, masked=True)
    stats = generate_abs_values_for_saving(True, data)
    (x_low, x_high, y_low, y_high, norm) = stats
    data1 = np.arange(x_low // 5 * 5, x_high // 5 * 5 + 5, 5)
    data2 = AbsAna.absorption_roi_masked[:, 1]
    data2 = np.clip(data2, a_min=y_low, a_max=y_high)
    name = get_save_abs_info(scale=True, image=False, masked=True, data=data, spec_number = spec_number)
    data = np.asarray([data1, data2]).T
    save_data(output_path, data, name, formatting="%.5f", gradient=True)
            
def save_absorption_spec_graph(output_path, data, spec_number, is_abspc_with_scale, is_abspc_wo_scale, masked, fmt=".png"):
    if is_abspc_with_scale:
        name = get_save_abs_info(scale=True, image=True, masked=masked, data=data, spec_number=spec_number)
        save_absorption_spec_diagram(output_path, data, name + "_0_derivative", True, masked, fmt=fmt)
        y_lim=[np.min(np.gradient(data)), np.max(np.gradient(data))]
        save_absorption_spec_diagram(output_path, np.gradient(data), name + "_1_derivative", True, masked, fmt=fmt, y_lim=y_lim)
        y_lim=[np.min(np.gradient(np.gradient(data))), np.max(np.gradient(np.gradient(data)))]
        save_absorption_spec_diagram(output_path, np.gradient(np.gradient(data)), name + "_2_derivative", True, masked, fmt=fmt, y_lim=y_lim)
    if is_abspc_wo_scale:
        name = get_save_abs_info(scale=False, image=True, masked=masked, data=data, spec_number = spec_number)
        save_absorption_spec_diagram(output_path, data, name + "_0_derivative", False, masked, fmt=fmt)
        y_lim=[np.min(np.gradient(data)), np.max(np.gradient(data))]
        save_absorption_spec_diagram(output_path, np.gradient(data), name + "_1_derivative", False, masked, fmt=fmt, y_lim=y_lim)
        y_lim=[np.min(np.gradient(np.gradient(data))), np.max(np.gradient(np.gradient(data)))]
        save_absorption_spec_diagram(output_path, np.gradient(np.gradient(data)), name + "_2_derivative", False, masked, fmt=fmt, y_lim=y_lim)
            
def save_absorption_spec_diagram(output_path, data, title, scale, masked, fmt=".png", y_lim = None):
    output_path = output_path + title + fmt
    logging.debug("SAVING ABSORPTION SPEC" + output_path)
    plt.clf()
    axes = plt.subplot(111)
    x_vals = np.arange(500, 1000, 5)
    stats = generate_abs_values_for_saving(masked, data)
    (x_low, x_high, y_low, y_high, norm) = stats
    if y_lim is not None:
        #title = title.replace(str(round(y_low,3)), str(round(y_lim[0], 4)))
        #title = title.replace(str(round(y_high,3)), str(round(y_lim[1], 4)))
        y_low = y_lim[0]
        y_high = y_lim[1]
        #output_path = self.current_output_path + "/" + title + fmt
    # plot absorption spec
    axes.plot(x_vals, data, '-', lw=0.5)
    axes.grid(linestyle=':', linewidth=0.5)
    low = y_low
    high = y_high
    if low is not None and high is not None:
        factor = (high - low) * 0.05
        low -= factor
        high += factor
    axes.set_xlim(left=x_low, right=x_high)
    axes.set_ylim(bottom=low, top=high)
    axes.ticklabel_format(style='plain')
    axes.get_yaxis().set_major_formatter(
        matplotlib.ticker.FuncFormatter(format_axis))
    axes.get_xaxis().set_major_formatter(
        matplotlib.ticker.FuncFormatter(format_axis))
    if scale:
        plt.title(title)
    else:
        plt.axis('off')
    if not os.path.exists(output_path):
        plt.savefig(output_path)
    plt.clf()
    
def format_axis(x, _):
    if x % 1 == 0:
        return format(int(x), ',')
    else:
        return format(round(x, 4))
    
def format_axis2(x, p):
    if x % 1 == 0:
        return format(int(x), ',')
    else:
        return format(round(x, 2))
    
def save_data(path, data, title, stats=[None, None], fmt=".csv", formatting="%.2f", gradient = False):
    output_path = path + title + fmt
    logging.debug("SAVING DATA TO " + output_path)
    if stats != [None, None]:
        data = np.clip(data, a_min=stats[0], a_max=stats[1])
    if gradient:
        x = data[:,0]
        f = np.round(data[:,1],5)
        f1 = np.round(np.gradient(f),5)
        f2 = np.round(np.gradient(f1),5)
        if not os.path.exists(path + title + ".xlsx"):
            workbook = xlsxwriter.Workbook(path + title + ".xlsx")
            row = 0
            row_2 = 0
            row_3 = 0
            col = 0
            worksheet = workbook.add_worksheet('0_derivative')
            worksheet2 = workbook.add_worksheet("1_derivative")
            worksheet3 = workbook.add_worksheet("2_derivative")
            for idx, i in enumerate(x):
                worksheet.write(row , col , i)
                worksheet.write(row , col + 1, f[idx])
                row += 1
            for idx, i in enumerate(x):
                worksheet2.write(row_2 , col , i)
                worksheet2.write(row_2, col + 1, f1[idx])
                row_2 += 1
            for idx, i in enumerate(x):
                worksheet3.write(row_3 , col , i)
                worksheet3.write(row_3, col + 1, f2[idx])
                row_3 += 1
            workbook.close()
        if not os.path.exists(path  + title + "_0_derivative" + fmt):
            np.savetxt(path +  title + "_0_derivative" + fmt, data, delimiter=",", fmt=formatting)
        if not os.path.exists(path + title + "_1_derivative" + fmt):
            np.savetxt(path + title + "_1_derivative" + fmt, np.array(np.transpose([x,f1])), delimiter=",", fmt=formatting)
        if not os.path.exists(path + title + "_2_derivative" + fmt):
            np.savetxt(path  + title + "_2_derivative" + fmt, np.array(np.transpose([x,f2])), delimiter=",", fmt=formatting)
    else:
        x = data[:,0]
        f = np.round(data[:,1],5)
        op = path + title + ".xlsx"
        if not os.path.exists(op):
            workbook = xlsxwriter.Workbook(op)
            row = 0
            col = 0
            worksheet = workbook.add_worksheet()
            for idx, i in enumerate(x):
                worksheet.write(row, col, i)
                worksheet.write(row, col+ 1, f[idx])
                row +=1
            workbook.close()
        op = path + title + fmt
        if not os.path.exists(op):
            np.savetxt(op, data, delimiter=",", fmt=formatting)
        
def generate_abs_values_for_saving(masked, data):
    x_vals = np.arange(500, 1000, 5)
    x_lower_scale_value = np.ma.min(x_vals)
    x_upper_scale_value = np.ma.max(x_vals)
    y_lower_scale_value = round(np.ma.min(data), 3)
    y_upper_scale_value = round(np.ma.max(data), 3)
    masked_stats = [x_lower_scale_value, x_upper_scale_value, y_lower_scale_value, y_upper_scale_value, False]
    arr = masked_stats
    if arr != [None, None, None, None, False]:
        a = [round(float(arr[i]), 4) for i in range(4)]
        a.append(arr[4])
        return a
    else:
        y_min = np.round(np.ma.min(data), 4)
        y_max = np.round(np.ma.max(data), 4)
        x_min = 500
        x_max = 995
        norm = False
        return [x_min, x_max, y_min, y_max, norm]
    
def get_save_abs_info(scale, image, masked, data, spec_number):
    # e.g. abspec_fromCSV1_(0-3)_(0-200000)_with-scale.png
    # e.g. abspec_fromCSV1_(0-3)_(0-200000)_with-scale_data.csv
    num = spec_number
    (xmin, xmax, ymin, ymax, normed) = generate_abs_values_for_saving(masked, data)
    norm = '_'
    if normed:
        norm = '-normed_'
    limits = '_(' + str(xmin) + '-' + str(xmax) + ')-(' + str(ymin) + '-' + str(ymax) + ')' + norm
    scale_mod = 'wo-scale'
    if scale:
        scale_mod = 'with-scale'
    masked_mod = 'whole_'
    if masked:
        masked_mod = 'masked_'
    if image:
        return 'spectrum_fromCSV' + str(num) + limits + masked_mod + scale_mod
    else:
        return 'spectrum_fromCSV' + str(num) + limits + masked_mod + 'data'

def read_dc(path):
    data = np.fromfile(path, dtype='>f')  # returns 1D array and reads file in big-endian binary format
    data_cube = data[3:].reshape(640, 480, 100)  # reshape to data cube and ignore first 3 values
    return data_cube   

def get_all_hypergui_folders(path):
    return glob.glob(path+'/**/_hypergui*/', recursive=True)
        
def hist_data_from_spec_num(HistAna):
    data = HistAna.histogram_data_masked.flatten()
    d = np.ma.sort(data)
    index = np.where(d.mask)[0][0]
    return d[:index]



def calc_stats(data):
    upper_value = np.ma.max(data.flatten())
    lower_value = np.ma.min(data.flatten())
    bins = np.arange(start=lower_value, stop=upper_value + 0.01,
                         step=0.01)
    
    
    histogram_data = np.histogram(data, bins=bins)
    # determine the minimum y value and at which x this occurs
    min_y = np.round(np.ma.min(np.histogram(data, bins=bins)[0]), 3)

    # determine the maximum y value and at which x this occurs
    max_y = np.round(np.ma.max(histogram_data[0]), 3)

    # determine the minimum x (bin) value and its size
    min_x = np.round(histogram_data[1][0], 3)

    # determine the maximum x (bin) value and its size
    max_x = np.round(histogram_data[1][-1], 3)
    
    return [min_x, max_x, min_y, max_y, 0.01]

    

def save_histogram(output_path, HistAna, spec_num):
    data = hist_data_from_spec_num(HistAna)
    save_histogram_graph(output_path, data, spec_num, True, True,
                                masked=True)
    data = hist_data_from_spec_num(HistAna)
    name = get_save_hist_info(scale=True, image=False, masked=True,
                                                data=data, spec_num = spec_num)
    save_histogram_data(output_path, data, name, masked=True)

def get_save_hist_info(scale, image, masked, data, spec_num):
    num = spec_num
    (xmin, xmax, ymin, ymax, step) = generate_hist_values_for_saving(masked, data)
    limits = '_(' + str(xmin) + '-' + str(xmax) + ')-(' + str(ymin) + '-' + str(ymax) + ')-' + str(step) + '_'
    scale_mod = 'wo-scale'
    if scale:
        scale_mod = 'with-scale'
    p_mod = ''
    data_mod = ['', '']
    masked_mod = 'whole_'
    if masked:
        masked_mod = 'masked_'
    if image:
        return 'histogram_fromCSV' + str(num) + limits + p_mod + data_mod[0] + masked_mod + scale_mod
    else:
        return 'histogram_fromCSV' + str(num) + limits + p_mod + data_mod[1] + masked_mod + 'data'

def generate_hist_values_for_saving(masked, data):
    arr = calc_stats(data)
    if arr != [None, None, None, None, None]:
        return [round(float(arr[i]), 4) for i in range(5)]
    else:
        lower = np.ma.min(data)
        upper = np.ma.max(data)
        step = 0.01
        bins = np.arange(start=lower, stop=upper + step, step=step)
        histogram_data = np.histogram(data, bins=bins)
        y_min = np.round(np.ma.min(histogram_data[0]), 4)
        y_max = np.round(np.ma.max(histogram_data[0]), 4)
        x_min = np.round(histogram_data[1][0], 4)
        x_max = np.round(histogram_data[1][-1], 4)
        return [x_min, x_max, y_min, y_max, step]
        
        
def save_histogram_data(output_path, data, name, masked):
    stats = generate_hist_values_for_saving(masked, data)
    (x_low, x_high, y_low, y_high, step) = stats
    start = x_low
    stop = x_high + step
    bins = np.arange(start=start, stop=stop, step=step)
    counts, hist_bins, _ = plt.hist(data, bins=bins)
    counts = np.clip(counts, a_min=y_low, a_max=y_high)
    hist_data = np.stack((bins[:-1], counts)).T
    save_data(output_path, hist_data, name, formatting="%.2f")

def save_histogram_graph(output_path, data, spec_num, is_hist_with_scale, is_hist_wo_scale, masked, fmt=".png"):
    if is_hist_with_scale:
        name = get_save_hist_info(scale=True, image=True, masked=masked,
                                                data=data, spec_num = spec_num)
        save_histogram_diagram(output_path, data, name, True, masked, fmt=fmt)
    if is_hist_wo_scale:
        name = get_save_hist_info(scale=False, image=True, masked=masked,
                                                data=data, spec_num = spec_num)
        save_histogram_diagram(output_path, data, name, False, masked, fmt=fmt)

def save_histogram_diagram(output_path, data, title, scale, masked, fmt=".png"):
    output_path = output_path + title + fmt
    logging.debug("SAVING HISTOGRAM TO " + output_path)
    plt.clf()
    axes = plt.subplot(111)
    stats = generate_hist_values_for_saving(masked, data)
    (x_low, x_high, y_low, y_high, step) = stats
    start = x_low
    stop = x_high + step
    bins = np.arange(start=start, stop=stop, step=step)
        # plot histogram
    axes.hist(data, bins=bins)
    axes.set_xlim(left=x_low, right=x_high)
    axes.set_ylim(bottom=y_low, top=y_high)
    # commas and non-scientific notation
    axes.ticklabel_format(style='plain')
    axes.get_yaxis().set_major_formatter(
        matplotlib.ticker.FuncFormatter(format_axis2))
    axes.get_xaxis().set_major_formatter(
        matplotlib.ticker.FuncFormatter(format_axis2)) 
    if scale:
        plt.title(title)
    else:
        plt.axis('off')
    if not os.path.exists(output_path):
        plt.savefig(output_path)
    plt.clf()
    
def save_mean_and_sd(path, dic):
    workbook = xlsxwriter.Workbook(path + "channel_numbers" + ".xlsx")
    col = 1
    worksheet = workbook.add_worksheet('channel_numbers')
    worksheet.write(0 , 1 , "STO2")
    worksheet.write(0 , 2 , "NIR")
    worksheet.write(0 , 3 , "THI")
    worksheet.write(0 , 4 , "TWI")
    worksheet.write(0 , 6 , "TLI")
    worksheet.write(0 , 5 , "OHI")
    worksheet.write(1 , 0 , "mean")
    worksheet.write(2 , 0 , "sd")
    worksheet.write(3 , 0 , "median")
    worksheet.write(4 , 0 , "IQR_low")
    worksheet.write(5 , 0 , "IQR_high")
    worksheet.write(6 , 0 , "min")
    worksheet.write(7 , 0 , "max")
    keys = ["sto2", "nir", "thi", "twi", "ohi", "tli"]
    for key in keys:
        tup = dic[key]
        if tup is not None:
            row = 1
            for val in tup:
                worksheet.write(row, col, val)
                row = row +1
        col = col +1
    workbook.close()

def get_and_save(path):
    if os.path.exists(path +'mask.csv'):
        mask = np.genfromtxt(path + 'mask.csv', delimiter=',')
        mask = np.fliplr(mask.T)
        mask = np.logical_not(mask)
        dc_path = glob.glob(os.path.dirname(os.path.dirname(path))+"/*SpecCube.dat")[0]
        data_cube = read_dc(dc_path)
        wavelength = (0 ,0)
             
        HistAna = HistogramAnalysis(dc_path, data_cube, wavelength, spec_tup5, mask)
        save_histogram(path, HistAna, 5)            
        AbsAna = AbsSpecAnalysis(data_cube, wavelength, spec_tup1, mask)
        save_absorption_spec(path, AbsAna, 1)
        
        HistAna = HistogramAnalysis(dc_path, data_cube, wavelength, spec_tup1, mask)
        save_histogram(path, HistAna, 1)
        AbsAna = AbsSpecAnalysis(data_cube, wavelength, spec_tup5, mask)
        save_absorption_spec(path, AbsAna, 5)
        
        mean_and_sd_channels = get_channels(mask, HistAna)
        save_mean_and_sd(path, mean_and_sd_channels)
        return True
    else:
        print("mask.csv does not exists in " + path)
        return False
    
def mask_to_tif(path):
    if len (glob.glob(path + "mask.csv"))>0:
        bin_mask = glob.glob(path + "mask.csv")[0]
        mask = np.genfromtxt(bin_mask, delimiter=',')
        mask_img = Image.fromarray(((mask*-1+1)*255).astype("uint8"), 'L')
        hg_name = os.path.basename(os.path.dirname(path))
        img_name = os.path.basename(os.path.dirname(os.path.dirname(path)))
        tif_name = path + img_name + "_mask" + hg_name + ".tif"
        if not os.path.exists(tif_name):
            mask_img.save(tif_name )
        else:
            print(tif_name + " exists. Skipping.")
      
    else:
        print("No binary mask in " + path)

jj = len(get_all_hypergui_folders(data_path))


ii = 0        
folderList = get_all_hypergui_folders(data_path)
for p in folderList:
    print('Current folder: ' + p)
    tb = glob.glob(p+ "_done_by_automated_hypergui_analysis_general.npy")
    if len(tb)==0:
        isDone = get_and_save(p)
        if isDone:
            np.save(p +"_done_by_automated_hypergui_analysis_general.npy", np.array([True]))
    ii = ii+1
    print(str(ii)+"/"+str(jj) + " folders complete.")
    
for hg in get_all_hypergui_folders(data_path):
    mask_to_tif(hg)
    
        
        
        