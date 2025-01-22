# Load Excel data
# Load Excel data

import pandas as pd
from pandas import read_excel


def load_excel_data(input_file_path):
    return read_excel(input_file_path, dtype={'NID': str})
