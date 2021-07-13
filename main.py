from datetime import datetime
import os
import shutil
import openpyxl
import pandas as pd

if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', None)

    today = datetime.today()
    dateTime = today.strftime("%m_%d_%y_%H%M%S")
