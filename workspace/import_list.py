import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from PIL import Image, ImageTk
import numpy as np
import paddlehub as hub
import cv2
import os
import pandas as pd
from lxml import etree
from openpyxl import *
import copy as copycopy
import io
import sys
import time
from ast import literal_eval
import pyautogui as pag

import global_value as gl

from test import results_demo

from OCR import OCR

from excel_all import write_to_excel_all, len_byte

from GUI import *

sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding='utf8') #改变默认输出的标准编码