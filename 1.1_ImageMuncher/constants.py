"""
@author: Alexander Studier-Fischer, Jan Odenthal, Berkin Özdemir, University of Heidelberg
"""
from pptx.util import Cm

#  If you want to change where the images are placed, you may change them here
#  If you mess up and forget what the original values are, you can find them in the repository

UPPER_ONE = (Cm(0.32), Cm(5.45))  # Format is (left, top)
UPPER_TWO = (Cm(6.59), Cm(5.45))
UPPER_THREE = (Cm(12.94), Cm(5.45))
UPPER_FOUR = (Cm(19.21), Cm(5.45))
MIDDLE_ONE = (Cm(0.32), Cm(9.96))
MIDDLE_TWO = (Cm(6.59), Cm(9.96))
MIDDLE_THREE = (Cm(12.94), Cm(9.96))
MIDDLE_FOUR = (Cm(19.21), Cm(9.96))
LOWER_ONE = (Cm(0.32), Cm(14.47))
LOWER_TWO = (Cm(6.59), Cm(14.47))
LOWER_THREE = (Cm(12.94), Cm(14.47))
LOWER_FOUR = (Cm(19.21), Cm(14.47))

MAIN_IMAGE_SIZE = (Cm(5.88), Cm(4.3))  # Format is (width, height)

#CAPTION_IMAGE_LOC = (Cm(1.26), Cm(5.84))  # The image left of the 2x2 square
CAPTION_IMAGE_SIZE = (Cm(5.57), Cm(4.07))

SLIDE_LAYOUT_TITLE_AND_CONTENT = 5

SHORT_TABLE_LOC = (Cm(1.26), Cm(3.18))
SHORT_TABLE_SIZE = (Cm(22.87), Cm(1))
SHORT_TABLE_ROW_NUM = 1
SHORT_TABLE_COL_NUM = 6

LONG_TABLE_LOC = (Cm(1.26), Cm(4.19))
LONG_TABLE_SIZE = (Cm(22.87), Cm(1))
LONG_TABLE_ROW_NUM = 1
LONG_TABLE_COL_NUM = 1

CAPTION_TXT_SIZE = (Cm(20), Cm(1.03))
CAPTION_TXT_LOC = (Cm(0), Cm(0))

RGB = "RGB"
OXY = "Oxygenation"
THI = "THI"
NIR = "NIR"
TWI = "TWI"
TLI = "TLI"
OHI = "OHI"

image_dic = {RGB: UPPER_ONE ,OXY: UPPER_TWO,NIR:UPPER_THREE ,THI: UPPER_FOUR,TWI: MIDDLE_ONE, OHI: MIDDLE_TWO, TLI: MIDDLE_THREE}