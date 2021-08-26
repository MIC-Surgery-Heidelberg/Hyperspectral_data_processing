# Hyperspectral_data_processing
#### Dependencies
imageio==2.9.0 <br>
matplotlib==3.4.3 <br>
numpy==1.19.4 <br>
pandas==1.0.3 <br>
Pillow==8.3.1 <br>
python-pptx==0.6.18 <br>
scikit-image==0.16.2 <br>
scipy==1.5.0 <br>
similaritymeasures==0.4.4 <br>
tensorflow==2.4.0 <br>
XlsxWriter==1.2.9 <br>


#### Installation
To run these Script, Python 3.7 is required (https://www.python.org/). For Installing dependencies, we recommend using pip (see https://packaging.python.org/tutorials/installing-packages/) <br>
To install all required dependencies, run `pip install imageio==2.9.0 matplotlib==3.4.3 numpy==1.19.4 pandas==1.0.3 Pillow==8.3.1 python-pptx==0.6.18 scikit-image==0.16.2 scipy==1.5.0 similaritymeasures==0.4.4 tensorflow==2.4.0 XlsxWriter==1.2.9`
#### Description
Descripe the repository
## 1.1_ImageMuncher
Description: To create a power point file with all the colour-coded index pictures from the TIVITA Hyperspectral System

## 1.2_XLSXtoPowerPoint
Description: To fill the table of the power point file with precise labels and descriptions

## 2.1_pptx_data-sort_and_delete
Description: To add all the primary HSI recordings to a data folder for extensive analysis

## 2.2_automated_hypergui_analysis
Description: To automatically process annotations done with the HyperGUI

## 2.3_crop_all_channels
Description: To crop all colour-coded index picutres to 640 x 480 pixles

## 2.4_mean_and_sd_extraction
Description: To extract spectral mean and standard deviation from processed annotations

## 2.5_spectrum_visualization
Description: To create a powerpoint visualizing the TIVITA recordings and respectively anotated regions

## 2.6_spectrum_listing
Description: To list all extracted spectra in a table

## 2.7_integral_and_heterogenity
Description: To calculate integral and heterogeneity of the spectral reflectance curves across different measurements
