import base64
import datetime as dt
import os
import pathlib
import pandas as pd
import random
import re
import sys


from dotenv import load_dotenv
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pandas import ExcelWriter
from sympy import symbols, latex
from tkinter import filedialog as fd

from core.constants import *
from core.exceptions import (
    FilePathException
)
from core.functionality.classes import (
    ProportionalInlineImage
)