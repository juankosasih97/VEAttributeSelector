from sheet import Sheet
import pandas as pd
from path_leaf import path_leaf
# The file.py contains a file class  which contains list of sheet instances and some variables which are used
# for maintaining easy access from viewer classes and summary of all dataframes from sheet instances in fl.

class File:
    def __init__(self):                 # constructor for file objects
        self.fl = list()
        self.baseindex = None           # stores index of baseline data sheet in fl
        self.baseline = None            # stores the baseline data
        self.bSumList = None            # stores the list of baseline sum data
        self.dfSumDiff = None           # stores the sum data of all dataframes in fl as summary (used during calculation and chart generation)
        self.file_path = None           # stores the sheet path of xlsx containing all the sheets of dataframes
        self.building_name = None       # stores the building name gotten from file_name
        self.selectedDFSumDiff = None   # stores user selected sum data

    def openFile(self, file_name):      # used to read an excel sheet and parses all its sheets contents to separate sheet objects in fl
        #sheet = Sheet(file_name)         # for separate excel sheet
        xls = pd.ExcelFile(file_name)   # initialize an excel sheet with path specified in file_name as xls
        self.file_path=file_name
        self.building_name, ext = path_leaf(file_name).split('.')
        for sheet_name in xls.sheet_names:      # loop for every sheets in xls parses its dataframe by calling sheet objects methods
            sheet = Sheet(self.building_name + ' - ' + sheet_name)           # add filename + " - " later
            sheet.zeroRemove(xls)
            if 'Max' not in sheet.df.index:
                sheet.addStatsRow()
            self.fl.append(sheet)

    def addDiffPercent(self):               # add to dfSumDiff the sum data of all other dataframes in flist as new row with its sheet_name as index
        for sheet in self.fl:
            if sheet.isBaseLine and not self.baseindex:         # does extra processing for files that is indicated as baseline(there can only be one baseline, else error)
                self.baseindex = self.fl.index(sheet)
                self.baseline = sheet.df
                self.bSumList = sheet.df.loc['Summed total']
                self.dfSumDiff = self.bSumList.to_frame().T
                self.dfSumDiff.rename({'Summed total': sheet.sheet_name}, axis = 'index', inplace = True)
        for sheet in self.fl:
            if self.fl.index(sheet) is not self.baseindex:
                if self.bSumList is not None:
                    seri = (sheet.df.loc['Summed total']-self.bSumList)/self.bSumList
                    seri = seri.to_frame().T
                    seri.rename({'Summed total': 'Diff'}, axis = 'index', inplace = True)
                    if 'Diff' not in sheet.df.index:
                        sheet.df = sheet.df.append(seri)
                    else:
                        sheet.df.loc[['Diff'],:]= seri
                self.dfSumDiff = self.dfSumDiff.append(sheet.df.loc['Summed total'].to_frame().T)
                self.dfSumDiff.rename({'Summed total': sheet.sheet_name}, axis='index', inplace=True)     #change sheet.sheet_name to path_leaf(sheet.sheet_name).split('_', 1)[1][:-5] for multiple files

    def clearFL(self):                      # reset the fl to default (empty)
        self.fl.clear()
        self.baseindex = None
        self.baseline = None
        self.bSumList = None
        self.file_path = None
        self.building_name = None
        self.selectedDFSumDiff = None
        self.dfSumDiff = None





