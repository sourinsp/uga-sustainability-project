import csv
import xlrd
import os
from os import listdir
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl.utils.cell import col
from openpyxl import load_workbook

sheet_2018 = "/Users/sourinpaturi/Dropbox/Projects/UGA Sustainability/UGAMain_2018.xls"
sheet_2019 = "/Users/sourinpaturi/Dropbox/Projects/UGA Sustainability/UGAMain_2019.xls"
sheet_2020 = "/Users/sourinpaturi/Dropbox/Projects/UGA Sustainability/UGAMain_2020.xls"

def main():

    df_2018 = pd.read_excel(sheet_2018)
    df_2019 = pd.read_excel(sheet_2019)
    df_2020 = pd.read_excel(sheet_2020)

    df_2018['Date'] = pd.to_datetime(df_2018['Date'], errors = 'coerce')
    df_2019['Date'] = pd.to_datetime(df_2019['Date'], errors='coerce')
    df_2020['Date'] = pd.to_datetime(df_2020['Date'], errors='coerce')


    algo_DST(df_2018, '2018')
    algo_DST(df_2019, '2019')
    algo_DST(df_2020, '2020')


def removeBreaks(df,year):

    if (year == "2018"):

        # Remove 2018 Breaks --------------

        # Remove Spring break

        spring_start = '2018-03-12'
        spring_end = '2018-03-16'

        df = df[(df['Date'] <= spring_start) | (df['Date'] > spring_end)]

        # Remove Summer break

        summer_start = '2018-04-25'
        summer_end = '2018-08-12'

        df = df[(df['Date'] < summer_start) | (df['Date'] > summer_end)]

        # Remove Fall breaks

        labor_day = '2018-09-03'
        fall_break = '2018-10-26'

        df = df[df['Date'] != labor_day]
        df = df[df['Date'] != fall_break]

        # Remove weekends

        df = df[~df['Day of the Week'].isin(['Saturday', 'Sunday'])]

        # Write to Excel

        df.to_excel("/Users/sourinpaturi/Dropbox/Projects/UGA Sustainability/UGA_DST_2018.xlsx")

    if (year == "2019"):

    # Remove 2019 Breaks ------------
        
        # Remove Spring break

        spring_start = '2019-03-11'
        spring_end = '2019-03-15'

        df = df[(df['Date'] <= spring_start) | (df['Date'] > spring_end)]


        # Remove Summer break

        summer_start = '2019-04-30'
        summer_end = '2019-08-15'

        df = df[(df['Date'] < summer_start) | (df['Date'] > summer_end)]

        # Remove Fall breaks

        labor_day = '2019-09-02'
        fall_break = '2019-11-01'

        df = df[df['Date'] != labor_day]
        df = df[df['Date'] != fall_break]

        # Remove weekends

        df = df[~df['Day of the Week'].isin(['Saturday', 'Sunday'])]

        # Write to Excel

        df.to_excel("/Users/sourinpaturi/Dropbox/Projects/UGA Sustainability/UGA_DST_2019.xlsx")



    if (year == "2020"):

    # Remove 2020 Breaks ------------

        # Remove Spring break

        spring_start = '2020-03-09'
        spring_end = '2020-03-13'

        df = df[(df['Date'] <= spring_start) | (df['Date'] > spring_end)]

        # Remove Summer break

        summer_start = '2020-04-28'
        summer_end = '2020-08-20'

        df = df[(df['Date'] < summer_start) | (df['Date'] > summer_end)]

        # Remove Fall breaks

        labor_day = '2020-09-07'
        fall_break = '2020-10-30'

        df = df[df['Date'] != labor_day]
        df = df[df['Date'] != fall_break]

        # Remove weekends

        df = df[~df['Day of the Week'].isin(['Saturday', 'Sunday'])]

        # Write to Excel

        df.to_excel("/Users/sourinpaturi/Dropbox/Projects/UGA Sustainability/UGA_DST_2020.xlsx")

def algo_DST(df, year):

    if (year == "2018"):

        start_date = '2018-03-11'
        end_date = '2018-11-04'

        df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

        df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')

        removeBreaks(df, '2018')

    if (year == "2019"):

        start_date = '2019-03-10'
        end_date = '2019-11-03'

        df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

        df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')
        
        removeBreaks(df, '2019')

    if (year == "2020"):

        start_date = '2020-03-08'
        end_date = '2020-11-01'

        df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
        
        df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')
        
        removeBreaks(df, '2020')

if __name__ == '__main__':
    main()