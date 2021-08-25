"""
@author: Alexander Studier-Fischer, Jan Odenthal, Berkin Özdemir, University of Heidelberg
"""

from tensorflow.python.keras.preprocessing.image import img_to_array, load_img
import glob
from PIL import Image
import os
from os.path import dirname, abspath
from pathlib import Path

base = os.getcwd().replace("\\", "/")


def getFolderlist(filepath):
    folderList=glob.glob(filepath + '/**/*.png', recursive = True)
    return folderList

def crop_images(filepath):
    folderList = getFolderlist(filepath)
    for png in folderList:
                    png = os.path.abspath(png)
                    if not glob.glob(os.path.abspath(Path(png).parent) + '/_crops') and not "_crops" in str(png) and not '_hypergui' in str(png):
                        os.mkdir(os.path.abspath(Path(png).parent) + '/_crops')
                    if "cropped" not in os.path.basename(png)[0:-4]:
                            if "TWI_segm.png" in os.path.basename(png) or "THI_segm.png" in os.path.basename(png) or "Perfusion_segm.png" in os.path.basename(png) or "Oxygenation_segm.png" in os.path.basename(png) or "RGB-Image" in os.path.basename(png)[0:-4] or "TLI_segm.png" in os.path.basename(png) or "OHI_segm.png" in os.path.basename(png):
                                if os.path.basename(png)[0:-4].endswith('r') != True and os.path.basename(png)[0:-4].endswith('e') != True:
                                    img_name=png
                                    if not os.path.exists(dirname(abspath(png)) + '/_crops/' + os.path.basename(img_name[0:-4])[0:19] + '_cropped' + os.path.basename(img_name[0:-4])[19:] +'.png'):
                                        others_image=load_img(img_name, color_mode='rgb')
                                        others_image_array=img_to_array(others_image)
                                        if others_image_array.shape[0] == 550:
                                            new_others_image_array=others_image_array[50:530,50:690,:]
                                        else:
                                            new_others_image_array=others_image_array[26:506,4:644,:]
                                        picture = Image.fromarray(new_others_image_array.astype('uint8'), 'RGB')
                                        picture.save(dirname(abspath(png)) + '/_crops/' + os.path.basename(img_name[0:-4])[0:19] + '_cropped' + os.path.basename(img_name[0:-4])[19:] +'.png')
                                        print('Image ' + os.path.basename(img_name[0:-4]) + ' cropped successfully!')
                                elif os.path.basename(png).endswith('RGB-Image.png') == True:
                                    img_name = png
                                    if not os.path.exists(dirname(abspath(png)) + '/_crops/' + os.path.basename(img_name[0:-4])[0:19] + '_cropped' + os.path.basename(img_name[0:-4])[19:] +'.png'):
                                        others_image = load_img(img_name, color_mode='rgb')
                                        others_image_array = img_to_array(others_image)
                                        if others_image_array.shape[0] == 550:
                                            new_others_image_array=others_image_array[50:530,20:660,:]
                                        else:
                                            new_others_image_array=others_image_array[30:510,3:643,:]
                                        picture = Image.fromarray(new_others_image_array.astype('uint8'), 'RGB')
                                        picture.save(dirname(abspath(png)) + '/_crops/' + os.path.basename(img_name[0:-4])[0:19] + '_cropped' + os.path.basename(img_name[0:-4])[19:] +'.png')
                                        print('Image ' + os.path.basename(img_name[0:-4]) + ' cropped successfully!')
                            if "TWI.png" in os.path.basename(png) or "THI.png" in os.path.basename(png) or "Perfusion.png" in os.path.basename(png) or "Oxygenation.png" in os.path.basename(png) or "RGB-Image" in os.path.basename(png)[0:-4] or "TLI.png" in os.path.basename(png) or "OHI.png" in os.path.basename(png):
                                if os.path.basename(png)[0:-4].endswith('r') != True and os.path.basename(png)[0:-4].endswith('e') != True:
                                    img_name=png
                                    if not os.path.exists(dirname(abspath(png)) + '/_crops/' + os.path.basename(img_name[0:-4])[0:19] + '_cropped' + os.path.basename(img_name[0:-4])[19:] +'.png'):
                                        others_image=load_img(img_name, color_mode='rgb')
                                        others_image_array=img_to_array(others_image)
                                        if others_image_array.shape[0] == 550:
                                            new_others_image_array=others_image_array[50:530,50:690,:]
                                        else:
                                            new_others_image_array=others_image_array[26:506,4:644,:]
                                        picture = Image.fromarray(new_others_image_array.astype('uint8'), 'RGB')
                                        picture.save(dirname(abspath(png)) + '/_crops/' + os.path.basename(img_name[0:-4])[0:19] + '_cropped' + os.path.basename(img_name[0:-4])[19:] +'.png')
                                        print('Image ' + os.path.basename(img_name[0:-4]) + ' cropped successfully!')
                                elif os.path.basename(png).endswith('RGB-Image.png') == True:
                                    img_name = png
                                    if not os.path.exists(dirname(abspath(png)) + '/_crops/' + os.path.basename(img_name[0:-4])[0:19] + '_cropped' + os.path.basename(img_name[0:-4])[19:] +'.png'):
                                        others_image = load_img(img_name, color_mode='rgb')
                                        others_image_array = img_to_array(others_image)
                                        if others_image_array.shape[0] == 550:
                                            new_others_image_array=others_image_array[50:530,20:660,:]
                                        else:
                                            new_others_image_array=others_image_array[30:510,3:643,:]
                                        picture = Image.fromarray(new_others_image_array.astype('uint8'), 'RGB')
                                        picture.save(dirname(abspath(png)) + '/_crops/' + os.path.basename(img_name[0:-4])[0:19] + '_cropped' + os.path.basename(img_name[0:-4])[19:] +'.png')
                                        print('Image ' + os.path.basename(img_name[0:-4]) + ' cropped successfully!')


crop_images(base)