import pandas as pd
import numpy as np

# This sheet.py contains a file class which holds a dataframe and its identity.
# This class acts as a model class in this system.'''

class Sheet:
    def __init__(self, sheet_name):       # constructor of sheet object
        self.sheet_name = sheet_name          # holds the dataframe name, taken from excel file sheet name (Baseline, daylightsensors, etc.)
        if 'baseline' in self.sheet_name: self.sheet_name.replace('baseline', 'Baseline')        # standardise the name for baseline data as it's pivotal to most calculations
        self.df = None                      # stores the entire table reads from excel sheet
        self.df_temp = None                 # stores the result table/dataframe with columns selected by user
        self.isBaseLine = False             # boolean value for maintaining identity of baseline data

    def zeroRemove(self, xls):      # this function reads and initialize the table from excel to dataframe before further processing, xls is file path of excel file

        #if sheet_name[-3:] == 'csv':
         #   df = pd.read_csv(sheet_name, index_col=0)
          #  if df.iloc[0].isnull().values.any() :
           #     df.drop(df.index[:2], inplace=True)         # This line is used for excel files copied from IESVE to remove empty and unimportant rows
            #df = df.loc[:, (df != '0').any(axis=0)]
            #return df
        #else:
            file_name, sheet_name = self.sheet_name.split(' - ')                             # get the sheet_name from the 'sheet_name'
            self.df = pd.read_excel(xls, sheet_name, index_col=0, header=None)              # read table from file specified in 'xls' on sheet 'sheet_name'
            self.df.columns = self.df.iloc[0]                                               # set the column for dataframe as the first row
            #if self.df.iloc[0].isnull().values.any() :
            self.df.drop(self.df.index[:3], inplace=True)                                   # removes junk lines from dataframe when extracted from IESVE
            self.df = self.df.loc[:, (self.df != 0).any(axis=0)]                            # removes all columns that contains all zeroes as value
            self.df = self.df.apply(pd.to_numeric)                                          # change all values to type numeric, generate exception if not possible cast to numeric
            self.df = self.df.select_dtypes(exclude=['object'])                             # removes non-numeric data
            self.df = self.df.groupby(self.df.columns, axis=1).sum()                        # sum the value of the dataframe groupedby column names (for room specific attribute)
            self.df = self.df * 1000                                                        # change the value from MWh to kWh
            for colNames in list(self.df):
                self.df.rename(columns={colNames : colNames.replace('MWh','kWh')}, inplace= True)   # change all 'MWh' in column names to 'kWh'
            if 'Baseline' in self.sheet_name: self.setBaseLine(True)                         # set a file as baseline if its sheet name contains 'Baseline'
            # commented code below are for calculating conduction gain if it exists
            #if ('Conduction gain - external windows (kWh)' in self.df.columns) \
            #    and ('Solar gain (kWh)' in self.df.columns) \
            #    and ('ApHVAC chillers load (kWh)' in self.df.columns):
            #    windowLoad = (self.df.loc[:, 'Solar gain (kWh)'] + self.df.loc[:,
            #                                                       'Conduction gain - external windows (kWh)']) \
            #                 / self.df.loc[:, 'ApHVAC chillers load (kWh)']
            #    self.df = self.df.assign(newData = windowLoad)
            #    self.df.rename(columns={'newData' : 'Window Cooling Load Conduction Gain (kWh)'}, inplace=True)
            #if ('Conduction gain - external walls (kWh)' in self.df.columns) \
            #    and ('Solar gain (kWh)' in self.df.columns) \
            #    and ('ApHVAC chillers load (kWh)' in self.df.columns):
            #    wallLoad = (self.df.loc[:, 'Solar gain (kWh)'] + self.df.loc[:,
            #                                                       'Conduction gain - external walls (kWh)']) \
            #                 / self.df.loc[:, 'ApHVAC chillers load (kWh)']
            #    self.df = self.df.assign(newData = wallLoad)
            #    self.df.rename(columns={'newData' : 'Wall Cooling Load Conduction Gain (kWh)'}, inplace=True)



    def addStatsRow(self):              # This function adds new indexes (max, min, mean, range) to the dataframe
        df_max = self.df.iloc[:12].max().to_frame().T
        df_max.rename({0: 'Max'}, axis ='index', inplace = True)
        df_min = self.df.min().to_frame().T
        df_min.rename({0: 'Min'}, axis ='index', inplace = True)
        df_mean = self.df.iloc[:12].mean().to_frame().T
        df_mean.rename({0: 'Mean'}, axis ='index', inplace = True)
        df_range = (self.df.iloc[:12].max()-self.df.min()).to_frame().T
        df_range.rename({0: 'Range'}, axis ='index', inplace = True)
        df1 = df_max.append(df_min, verify_integrity = True)
        df2 = df_mean.append(df_range, verify_integrity = True)
        df1 = df1.append(df2, verify_integrity = True)
        self.df = self.df.append(df1, verify_integrity = True)

    def setBaseLine(self, bool):        # setter function for isBaseline class variable(also possible to access directly)
        self.isBaseLine = bool