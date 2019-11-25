from PyQt5 import QtCore, QtWidgets, QtGui
import pandas as pd
from pandas_model import PandasModel
from cList import cList
from file import File
from chartSelect import Ui_Dialog
from path_leaf import path_leaf
from financeCalc import Ui_InvestmentDialog
from categoryInterface import Ui_categoryDialog
from exceptionHandler import excepthook
import sys

# This module is the main window as view as well as the controller of the whole application

class Ui_MainWindow(object):

    def setupUi(self, MainWindow):          # setting up all the widgets in the main window as well as class variables
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(814, 726)


        self.flist = list()                   #File()      # create a new list of file list
        self.clist = cList()            # create a new category list
        #if not self.clist.catFileReadable:
        #    junk = QtWidgets.QMessageBox.warning(self.centralwidget, 'Category data reading failed.',
        #                                         'The file savecat.txt cannot be read, '
        #                                         'using default category setting instead.')
        self.clUI = list()              # create a list for category shown in the window
        self.selectedDF = None          # copy the dataframe being selected to this variable for showing in tableView
        self.selectedDFSumDiff = None   # maintain the summary dataframe with only the selected attributes
        self.index = 0                  # maintain the index(from filelist) of the file being shown in tableView
        self.current_buildingI = None   # maintain the index of the building(file) data being shown in tableView
        self.current_techI = None       # maintain the index of the technology(sheet) data being shown in tableView
        self.dialog = None              # initialize a class variable for storing chartDialog
        self.finLog = None              # initialize a class variable for storing investmentDialog
        self.allAttList = list()        # maintain a list of all attributes available in any table
        self.attList = list()           # maintain a list of all user-selected attributes
        self.chartList = list()         # a list of chart specification that is retrieved from chartDialog
        self.finData = None             # a dataframe for financial calculation in investmentDialog
        self.defaultcat = True



        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.gridLayout_2 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_2.setObjectName("gridLayout_2")

        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")

        self.allAttributes = QtWidgets.QTreeWidget(self.centralwidget)              # QTreeWidget that handles showing all attributes from input xlsx based on
        self.allAttributes.setObjectName("allAttributes")                           # categories
        self.allAttributes.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)  # enables user to select more than one item
        self.allAttributes.itemDoubleClicked.connect(self.add)          #  connect doubleclicked signal to add function
        self.load = QtWidgets.QTreeWidgetItem(self.allAttributes)           # create an item for general load
        self.energy = QtWidgets.QTreeWidgetItem(self.allAttributes)         # create an item for general energy
        self.breakdown = QtWidgets.QTreeWidgetItem(self.allAttributes)      # create an item for other attributes
        self.addCList()                                                     # call addCList which adds more item per category

        self.verticalLayout_2.addWidget(self.allAttributes)

        self.gridLayout_2.addLayout(self.verticalLayout_2, 0, 0, 1, 1)

        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")

        self.selectedAttributes = QtWidgets.QListWidget(self.centralwidget)         # QListWidget showing all item added by user from allAttributes
        self.selectedAttributes.setDragEnabled(True)
        self.selectedAttributes.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)       # enable drag and drop to change order of item in the list
        self.selectedAttributes.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.selectedAttributes.setObjectName("selectedAttributes")
        self.selectedAttributes.itemPressed.connect(self.editTable)                 # connect itemPressed signal to editTable to change order of attributes

        self.verticalLayout_3.addWidget(self.selectedAttributes)

        self.gridLayout_2.addLayout(self.verticalLayout_3, 0, 2, 1, 1)


        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")

        self.addButton = QtWidgets.QPushButton(self.centralwidget)              # add button to add attributes selected from allAttributes to selectedAttributes
        self.addButton.setObjectName("addButton")
        self.addButton.clicked.connect(self.allAttSelectionWrapper)             # connect click to function for adding attributes

        self.verticalLayout_4.addWidget(self.addButton)

        self.removeButton = QtWidgets.QPushButton(self.centralwidget)           # remove button to remove attributes selected from selectedAttributes
        self.removeButton.setObjectName("removeButton")
        self.removeButton.clicked.connect(self.remove)

        self.verticalLayout_4.addWidget(self.removeButton)

        self.chartButton = QtWidgets.QPushButton(self.centralwidget)            # chart button to call the chartDialog
        self.chartButton.setObjectName('chartButton')
        self.chartButton.clicked.connect(self.openChartDialog)

        self.verticalLayout_4.addWidget(self.chartButton)

        self.finButton = QtWidgets.QPushButton(self.centralwidget)              # finance button to call investmentDialog
        self.finButton.setObjectName('finButton')
        self.finButton.clicked.connect(self.openFinDialog)

        self.verticalLayout_4.addWidget(self.finButton)

        self.gridLayout_2.addLayout(self.verticalLayout_4, 0, 1, 1, 1)

        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")

        #self.checkbox1 = QtWidgets.QCheckBox(self.centralwidget)
        #self.checkbox1.setText('Calculate window cooling load conduction gain')
        #self.checkbox1.clicked.connect(lambda:self.btnstate(self.checkbox1))

        #self.gridLayout.addWidget(self.checkbox1, 0, 0, 1, 5)

        #self.checkbox2 = QtWidgets.QCheckBox(self.centralwidget)
        #self.checkbox2.setText('Calculate wall cooling load conduction gain')
        #self.checkbox2.clicked.connect(lambda:self.btnstate(self.checkbox2))

        #self.gridLayout.addWidget(self.checkbox2, 0, 6, 1, 5)

        self.buttonBox = QtWidgets.QDialogButtonBox(self.centralwidget)         # a set of standard button in the interface
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Open|QtWidgets.QDialogButtonBox.Reset|QtWidgets.QDialogButtonBox.Save) #open, reset, and save exists
        self.buttonBox.setObjectName("buttonBox")
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Open).clicked.connect(self.openFileWrapper)    # click open to open input xlsx file
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Save).clicked.connect(self.saveFile)           # click save to save formatting to an xlsx file
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Reset).clicked.connect(self.resetAllWrapper)   # click reset to reset the application

        self.gridLayout.addWidget(self.buttonBox, 1, 7, 1, 4)

        self.tablePicker = QtWidgets.QComboBox(self.centralwidget)          # a dropdown list to select a table or dataframe to be shown in tableView
        self.tablePicker.setObjectName("tablePicker")
        self.tablePicker.activated.connect(self.changeIndex)

        self.gridLayout.addWidget(self.tablePicker, 1, 0, 1, 5)

        self.catButton = QtWidgets.QPushButton(self.centralwidget)         # button to show a change unit dialog (not yet implemented)
        self.catButton.setObjectName('catButton')
        self.catButton.setText('Edit Category')
        self.catButton.clicked.connect(self.openCatDialog)

        self.gridLayout.addWidget(self.catButton, 1, 6, 1, 1)

        self.tableViewer = QtWidgets.QTableView(self.centralwidget)         # tableView to show dataframe data selected
        self.tableViewer.setObjectName("tableViewer")

        self.gridLayout.addWidget(self.tableViewer, 2, 0, 1, 11)

        self.gridLayout_2.addLayout(self.gridLayout, 1, 0, 1, 3)

        MainWindow.setCentralWidget(self.centralwidget)

        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 814, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)

        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "VE Attribute Selector"))
        self.allAttributes.headerItem().setText(0, _translate("MainWindow", "Attributes"))
        __sortingEnabled = self.allAttributes.isSortingEnabled()
        self.allAttributes.setSortingEnabled(False)
        self.allAttributes.topLevelItem(0).setText(0, _translate("MainWindow", "General Load"))
        self.allAttributes.topLevelItem(1).setText(0, _translate("MainWindow", "General Energy"))
        self.allAttributes.topLevelItem(2).setText(0, _translate("MainWindow", "Other"))
        self.allAttributes.setSortingEnabled(__sortingEnabled)
        self.addButton.setText(_translate("MainWindow", "Add"))
        self.removeButton.setText(_translate("MainWindow", "Remove"))
        self.chartButton.setText(_translate('MainWindow', 'Charts'))
        self.finButton.setText(_translate('MainWindow', 'Finance'))


    def openFileWrapper(self):      #  function to call openfile and initialize the insertion of attributes to the QTreeWidget
        self.openfile()
        attList = list()
        for file in self.flist:
            for sheet in file.fl:
                attributeList = sheet.df.columns.values.tolist()
                attList.extend(attributeList)
                attList = list(set(attList))
                for x in attributeList:
                    if x != 'Window Cooling Load Conduction Gain (kWh)' and x != 'Wall Cooling Load Conduction Gain (kWh)':
                        self.insertTree(x)
                        self.checkCList(x)
        self.allAttList = attList
        for file in self.flist:
            for sheet in file.fl:
                for att in attList:                                    # add column with all zeroes if att not in dataframe
                    if att not in sheet.df.columns:
                        sheet.df = sheet.df.assign(att=0)
                        sheet.df.rename({'att': att}, axis = 'columns', inplace = True)
        for file in self.flist:
            file.addDiffPercent()


    def insertTree(self, x):            # add new items from xlsx to QTreeWidget based on different categories (general load, general energy, others)
        if str(x)[-10:-6] == 'load' or str(x)[-11:-6] == 'input' or str(x)[-10:-6] == 'heat' or str(x)[-12:-6] == 'demand':
            for n in range(self.load.childCount()):
                if self.load.child(n).text(0) == x: return  # does not add the item if it exists in the QTreeWidget
            item = QtWidgets.QTreeWidgetItem(self.load)     # add item to load if attribute name pass condition above
            item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsDragEnabled | QtCore.Qt.ItemIsDropEnabled | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            _translate = QtCore.QCoreApplication.translate
            item.setText(0, _translate('MainWindow', str(x)))
        elif str(x)[-12:-6] == 'energy' or str(x)[:5] == 'Total' or str(x)[:5] == 'Equip' or str(x)[:6] == 'Lights' or str(x)[:6] == 'System':
            for n in range(self.energy.childCount()):
                if self.energy.child(n).text(0) == x: return
            item = QtWidgets.QTreeWidgetItem(self.energy)
            item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsDragEnabled | QtCore.Qt.ItemIsDropEnabled | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            _translate = QtCore.QCoreApplication.translate
            item.setText(0, _translate('MainWindow', str(x)))
        else:
            for n in range(self.breakdown.childCount()):
                if self.breakdown.child(n).text(0) == x: return
            item = QtWidgets.QTreeWidgetItem(self.breakdown)
            item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsDragEnabled | QtCore.Qt.ItemIsDropEnabled | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            _translate = QtCore.QCoreApplication.translate
            item.setText(0, _translate("MainWindow", str(x)))

    def openfile(self):             # function to call a QFileDialog to open an xlsx file
        fdlg = QtWidgets.QFileDialog()
        fdlg.setFileMode(QtWidgets.QFileDialog.ExistingFiles)        # only enable user to select one existing file
        fdlg.setNameFilter('Excel file(*.xlsx)')                    # only enable to open xlsx file
        if fdlg.exec_():                                            #
            for file in self.flist:
                for openfile in fdlg.selectedFiles():
                    if file.file_path == openfile:         # test if file is already open, warning given
                        junk = QtWidgets.QMessageBox.warning(self.centralwidget, 'File opened',
                                                               'The selected file ' + path_leaf(openfile) + ' is already opened, '
                                                               'select a different file for formatting')
                        return None
                # if self.flist.file_path is not None:                            # check if user selected file, warns user with QMessageBox about losing formatting
                    # answer = QtWidgets.QMessageBox.question(self.centralwidget, 'Save before proceeding', 'Any unsaved formatting'
                                                                                            # ' will be lost. Proceed to open'
                                                                                            # ' a different file?')
                    # if answer != QtWidgets.QMessageBox.Yes:
                        # return None
                # self.resetAll()                 # reset everything before parsing the new file
            for x in fdlg.selectedFiles():
                file_name = x
                if file_name:
                    file_name = str(file_name)
                        #self.tablePicker.addItem(file_name)
                    f = File()
                    try:
                        f.openFile(file_name)          # call the filelist method to open the file (reads separate sheets as different 'file'
                        for sheet in f.fl:
                            self.tablePicker.addItem(sheet.sheet_name)  # add each 'file' to the dropdown list
                        # f.addDiffPercent()                      # call the file method to summarise the data
                        self.flist.append(f)
                    except:
                        junk = QtWidgets.QMessageBox.warning(self.centralwidget, 'File cannot be opened',
                                                             'The selected file ' + path_leaf(
                                                                 file_name) + ' cannot be parsed, '
                                                                             'check if position of table and formatting is proper')
                        continue
            fdlg.close()
        return None


    def add(self, item):            # checks if selected item is a category
        if self.flist:
            if item.text(0) == 'General Load' or item.text(0) == 'General Energy' or \
                    item.text(0) == 'Other':                            # if general category, skip (have not implemented insertion using this)
                return
            for cat in self.clist.cl:
                if cat.name == item.text(0):                            # if specific category, add all of its members/child that exists in allAttributes
                    for member in cat.members:
                        for file in self.flist:
                            if member in list(file.dfSumDiff.columns.values) and self.allAttributes.findItems(member, QtCore.Qt.MatchRecursive|QtCore.Qt.MatchExactly): self.addItem(member)
                    return
            self.addItem(item.text(0))                                  # if not category then add normally by addItem


    def addItem(self, stritem):                                         # add attributes from QTreeWidget
        for x in range(self.selectedAttributes.count()):
            if self.selectedAttributes.item(x).text() == stritem:       # if item is in selectedAttributes, then don't add anymore
                return
        item = QtWidgets.QListWidgetItem()
        item.setFlags(
            QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsDragEnabled | QtCore.Qt.ItemIsDropEnabled | QtCore.Qt.ItemIsEnabled)
        self.selectedAttributes.addItem(item)                           # else add the item to selectedAttributes
        _translate = QtCore.QCoreApplication.translate
        item.setText(_translate("MainWindow", stritem))
        self.editTable()                                                # call editTable to process the change on the table
        if not self.dialog:                                             # if chartDialog haven't been initialised yet, initialise it.
            self.dialog = QtWidgets.QDialog()
            self.dialog.ui = Ui_Dialog()
            self.dialog.ui.setupUi(self.dialog, self.flist, self.clist, self.attList)   # pass the filelist, category list, and the selected attribute list to chartDialog
            # self.dialog.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        else:
            self.dialog.ui.addEntry(self.flist, self.clist, self.attList)   # update new or removed entries to chartDialog by passing updated data
        if not self.chartList:                                          # if chartList is empty, update the chartList with data from selectedCharts in chartDialog
            for x in range(self.dialog.ui.selectedCharts.count()):
                self.chartList.append(self.dialog.ui.selectedCharts.item(x).text())


    def remove(self):           # remove selected items from selectedAttributes list
        if self.flist:
            if self.selectedDF is not None:            # only removes anything if any data exists
                selectedSelAt = self.selectedAttributes.selectedItems()
                for item in selectedSelAt:
                    self.selectedAttributes.takeItem(self.selectedAttributes.row(item)) #remove selectedItems from QListWidget
                self.editTable()                # edit the change on the table accordingly
                self.dialog.ui.addEntry(self.flist, self.clist, self.attList)       # edit the change on chartDialog
                for item in selectedSelAt:
                    indexList = list()
                    for index in range(self.dialog.ui.selectedCharts.count()):
                        if item.text() in self.dialog.ui.selectedCharts.item(index).text():
                            indexList.append(index)         # note all the index of selectedChart item which uses attributes that are to be removed
                    for index in reversed(indexList):       # remove the item in reverse order to bypass the problem of changing index
                        self.dialog.ui.selectedCharts.takeItem(index)
                self.chartList.clear()                      # reupdate the chartList with the existing items
                for x in range(self.dialog.ui.selectedCharts.count()):
                    self.chartList.append(self.dialog.ui.selectedCharts.item(x).text())


    def allAttSelectionWrapper(self):           # a wrapper for item adding
        selectedAllAt = self.allAttributes.selectedItems()
        for item in selectedAllAt:
            self.add(item)

    def changeIndex(self, index):               # called when user activates drop down list (changing technology being shown on tableView)
        self.index = index
        str = self.tablePicker.itemText(index)
        current_building, current_tech = str.split(' - ')
        for index, file in enumerate(self.flist):
            if file.building_name == current_building:
                self.current_buildingI = index
                for ind, sheet in enumerate(self.flist[index].fl):
                    if current_tech in sheet.sheet_name: self.current_techI = ind
        self.editTable()


    def editTable(self):                        # maintains the change in attList, each file's df_temp (containing selected attributes only), and selected Data
        self.attList.clear()
        for x in range(self.selectedAttributes.count()):
            self.attList.append(self.selectedAttributes.item(x).text())     # update the attList according to items in selectedAttributes
        for file in self.flist:
            for sheet in file.fl:
                for att in self.attList:                                    # add column with all zeroes if att not in dataframe
                    if att not in sheet.df.columns:
                        sheet.df = sheet.df.assign(att=0)
                        sheet.df.rename({'att': att}, axis = 'columns', inplace = True)
                sheet.df_temp = sheet.df.filter(self.attList)                           # store a temporary dataframe with all selected Attributes for each 'file'(technology)
        stri = self.tablePicker.itemText(self.index)
        current_building, current_tech = stri.split(' - ')
        for index, file in enumerate(self.flist):
            if file.building_name == current_building:
                self.current_buildingI = index
                for ind, sheet in enumerate(self.flist[index].fl):
                    if current_tech in sheet.sheet_name: self.current_techI = ind
        self.selectedDF = self.flist[self.current_buildingI].fl[self.current_techI].df_temp                 # store the selected temporary dataframe based on dropdownlist / index
        for file in self.flist:
            if file.dfSumDiff is not None:
                self.selectedDFSumDiff = file.dfSumDiff.filter(self.attList)  # store a summary dataframe with only the selected Attributes
                file.selectedDFSumDiff = file.dfSumDiff.filter(self.attList)
        model = PandasModel(self.selectedDF)
        self.tableViewer.setModel(model)                    # create a model of table using the selected temporary dataframe and show it in tableView


    def saveFile(self):                             # save the formatting to user-selected xlsx file, complete with all charts and financial data
        if self.flist and self.selectedDF is not None:          # only save if user has selected attributes
            sfdlg = QtWidgets.QFileDialog()
            sfdlg.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)       # call a QFileDialog which saves only xlsx file
            sfdlg.setNameFilter('Excel file(*.xlsx)')
            if sfdlg.exec_():
                file_name = sfdlg.selectedFiles()[0]
                #if file_name[-3:] == 'csv':
                 #   self.df_temp.to_csv()
                #else :
                count = 0
                self.writer = pd.ExcelWriter(file_name, engine='xlsxwriter')        # store a writer to the variable
                self.workbook = self.writer.book                                    # store the book to the variable
                self.format1 = self.workbook.add_format({'num_format': '#,###,##0.0##'})    # store a format for decimal value and thousand-separator
                self.formatpercent = self.workbook.add_format({'num_format': '0.0##%'})     # store a format for percentage
                rowcount = 0
                ##these codes below are for including general charts in excel files, special user specified charts code are below this section
                for cat in self.clist.cl:
                    df = pd.DataFrame()
                    rowcount = 0
                    strcol = 0
                    for file in self.flist:
                        if file.selectedDFSumDiff is not None:  # only save chart if summary data exists
                            df_temp = file.selectedDFSumDiff[file.selectedDFSumDiff.columns.intersection(cat.members)]      # select only the existing member of the category from summary data
                            strcol = len(df_temp.columns) + 4
                            if self.defaultcat:
                                if cat.name != 'Cooling load':
                                    df_temp = df_temp.loc[:,[col for col in df_temp.columns if 'DX' not in col]]
                                else: df_temp = df_temp.loc[:, ['Ap Sys chillers load (kWh)', 'ApHVAC chillers load (kWh)']]
                            #if cat.name == 'Building energy':
                            #    acmv = self.clist.getMember('Chiller energy')
                            #   acmvDF = self.selectedDFSumDiff[self.selectedDFSumDiff.columns.intersection(acmv)]
                            #    acmvTotal = acmvDF.sum(axis = 1)
                            #    df_temp = df_temp.assign(ACMV=acmvTotal)                                                    # add acmv(chiller energy) data to building energy
                            if len(df_temp.columns):
                                df_temp = df_temp.T.sort_values(file.building_name+' - Baseline', ascending=False).T        # sort the values by column in descending order
                                if self.defaultcat:
                                    if cat.name != 'Cooling load':      # special case for cooling load, remove DX for others
                                        df_temp = df_temp.T.append(df_temp.loc[:,[col for col in df_temp.columns if 'DX' not in col]].sum(axis='columns').to_frame().T)        # add the sum of the category as a new column
                                    else: df_temp = df_temp.T.append(df_temp.loc[:, ['Ap Sys chillers load (kWh)', 'ApHVAC chillers load (kWh)']].sum(axis='columns').to_frame().T) # exclude total chiller load from sum
                                else:
                                    df_temp = df_temp.T.append(df_temp.sum(axis='columns').to_frame().T)        # unless category is default, sum everything available in table
                                df_temp.rename({0: 'Sum'}, axis='index', inplace=True)
                                df_temp = df_temp.T
                                df_temp.reset_index(inplace=True)
                                df_temp.loc[1:]=df_temp[1:].sort_values('Sum').values       # sort the sum values in ascending order leaving baseline as first row
                                df_temp.set_index(df_temp.columns.values.tolist()[0], inplace=True)
                                #strcol =  len(file.selectedDFSumDiff.columns) + 3          # find the length of dataframe to position the table in excel
                                df_temp.to_excel(self.writer, sheet_name=cat.name, startrow = rowcount, startcol=strcol)     # write the dataframe to excel on sheet: category name

                                worksheet = self.writer.sheets[cat.name]            # store the worksheet for adding chart
                                chart1 = self.workbook.add_chart({'type': 'column'})        # add a new column chart/bar chart
                                chart1.add_series({'values': [cat.name, 1 + rowcount, strcol + len(df_temp.columns), rowcount + len(file.fl) ,  # use the value on the first row of the table
                                                              strcol + len(df_temp.columns)],
                                                   'categories': [cat.name, rowcount + 1, strcol, rowcount + len(file.fl), strcol],     # use the attribute names as category(x-axis)
                                                   'data_labels': {'value': True},                                      # show the data label on chart
                                                   'name': cat.name  +' ('+ file.building_name +')'})
                                chart1.set_y_axis({'min': 0, 'name': 'kWh'})
                                chart1.set_legend({'position': 'none'})
                                worksheet.insert_chart(rowcount , strcol + len(df_temp.columns) + 2, chart1)                   # insert the chart based on the length of the table
                                if len(df_temp.columns) > 2:                                                            # add a pie chart if multiple attributes exists in the category
                                    index = df_temp.index.get_loc(df_temp.index.values.tolist()[0])+1                   # to show how much each attributes takes up part in the sum
                                    chartpie = self.workbook.add_chart({'type': 'pie'})
                                    chartpie.add_series({'values': [cat.name, rowcount + index, strcol + 1, rowcount + index,
                                                                    strcol + len(df_temp.columns) - 1],
                                                         'categories': [cat.name, rowcount, strcol + 1, rowcount,
                                                                        strcol + len(df_temp.columns) - 1],
                                                         'name': cat.name + ' breakdown ('+ file.building_name +')',
                                                         'data_labels': {'percentage': True}
                                                         })
                                    worksheet.insert_chart(rowcount, strcol + len(df_temp.columns) + 10, chartpie)
                                df = df.append(df_temp)         # add df_temp data to df for compilation to calculate and generate saving data
                        rowcount += file.selectedDFSumDiff.shape[0]+2
                    if df.columns.values.tolist():
                        self.saveSavingData(df, cat.name, strcol + df.shape[0] + 2)
                ##end of general charts code
                ##output dataframe to xlsx based on category
                rowcap = rowcount   # store the length(height) of the table for placing other tables and charts
                rowcount = 0
                for file in self.flist:
                    for f in file.fl:
                        f.df_temp.to_excel(self.writer, sheet_name='General', startrow=rowcount)       # add all selected data to a general sheet
                        worksheet = self.writer.sheets['General']
                        worksheet.set_column(1, len(f.df_temp.columns), None, self.format1) # set numeric format to the data
                        #f.df_temp.to_excel(writer, 'Sheet'+str(count))
                        for cat in self.clist.cl:
                            df_temp = f.df_temp[f.df_temp.columns.intersection(cat.members)]
                            if len(df_temp.columns):
                                df_temp.to_excel(self.writer, sheet_name=cat.name, startrow=rowcount )    # add attributes/columns to xlsx based on category
                                self.writer.sheets[cat.name].write(rowcount, 0, f.sheet_name) # add the tech name as table name #change f.sheet_name to path_leaf(f.sheet_name)
                                worksheet = self.writer.sheets[cat.name]
                                worksheet.set_column(1, len(df_temp.columns), None, self.format1)   # set numeric format to the data
                        rowcount+=f.df_temp.shape[0]+2
                ##specialized charts code
                countList = [0]*len(self.clist.cl)      # initialise a counterList for inserting specific charts to each categorical sheets
                genCount = 0                            # initialise a counter for inserting chart to 'general' sheet
                for entry in self.chartList:  # process each entry in the chartList obtained from chartDialog
                    file1, file2, attr = entry.split(', ')  # find the specification of the chart
                    current_building1, current_tech1 = file1.split(' - ')
                    current_buildingI1, current_buildingI2, current_techI1, current_techI2 = 0, 0, 0, 0
                    for index, file in enumerate(self.flist):
                        if file.building_name == current_building1:
                            current_buildingI1 = index
                    current_building2, current_tech2 = file2.split(' - ')
                    for index, file in enumerate(self.flist):
                        if file.building_name == current_building2:
                            current_buildingI2 = index
                    if attr == 'Sum':                   # add 'sum' chart for all category lists if attribute is sum
                        for cat in self.clist.cl:
                            df_temp1 = self.flist[current_buildingI1].selectedDFSumDiff[self.flist[current_buildingI1].selectedDFSumDiff.columns.intersection(cat.members)].loc[
                                        [file1], :]    # select the data based on the specification  file1
                            df_temp = df_temp1.append(self.flist[current_buildingI2].selectedDFSumDiff[self.flist[current_buildingI2].selectedDFSumDiff.columns.intersection(cat.members)].loc[
                                        [file2], :])
                            #if cat.name == 'Building energy':
                            #    acmv = self.clist.getMember('Chiller energy')
                            #    acmvDF = self.selectedDFSumDiff[self.selectedDFSumDiff.columns.intersection(acmv)]
                            #    acmvTotal = acmvDF.sum(axis=1)
                            #    df_temp = df_temp.assign(ACMV=acmvTotal)
                            if len(df_temp.columns):        # check if data exists
                                if self.defaultcat:
                                    if cat.name != 'Cooling load':      # special case for cooling load, remove DX for others
                                        df_temp = df_temp.T.append(df_temp.loc[:,[col for col in df_temp.columns if 'DX' not in col]].sum(axis='columns').to_frame().T)        # add the sum of the category as a new column
                                    else: df_temp = df_temp.T.append(df_temp.loc[:, ['Ap Sys chillers load (kWh)', 'ApHVAC chillers load (kWh)']].sum(axis='columns').to_frame().T) # exclude total chiller load from sum
                                else:
                                    df_temp = df_temp.T.append(df_temp.sum(axis='columns').to_frame().T)        # unless category is default, sum everything available in table
                                df_temp.rename({0: 'Sum'}, axis='index', inplace=True)
                                df_temp = df_temp.T             # add the sum data as column
                                df_length = len(df_temp.columns)        # find the length of the table for positioning
                                count = countList[self.clist.cl.index(cat)] # get the count for the specific category from countList for positioning
                                self.insertChart(file1, file2, attr, cat.name, count, df_temp, df_length-1, rowcap) # call insertChart
                                countList[self.clist.cl.index(cat)]+=1          # increment the countList by 1
                    else:                               # if attribute is not 'sum' then add the attribute chart based on the category
                        inCat = False                   # boolean to check if attribute is a member of specific category
                        df_temp1 = self.flist[current_buildingI1].selectedDFSumDiff[
                                       self.flist[current_buildingI1].selectedDFSumDiff.columns.intersection(
                                           [attr])].loc[
                                   [file1], :]  # select the data based on the specification  file1
                        df_temp = df_temp1.append(self.flist[current_buildingI2].selectedDFSumDiff[self.flist[
                            current_buildingI2].selectedDFSumDiff.columns.intersection([attr])].loc[
                                                  [file2], :])
                        for cat in self.clist.cl:
                            if attr in cat.members:     # add the chart to the category sheet if it is the member of the category
                                df_length = len(self.selectedDFSumDiff.columns.intersection(cat.members))
                                count = countList[self.clist.cl.index(cat)]     # check the count of chart that exists in the sheet
                                self.insertChart(file1, file2, attr, cat.name, count, df_temp, df_length, rowcap)
                                countList[self.clist.cl.index(cat)]+=1
                                inCat = True
                        if not inCat:                   # if not, then add it to the 'general' sheet
                            df_length = len(self.selectedDFSumDiff.columns)
                            self.insertChart(file1, file2, attr, 'General', genCount, df_temp, df_length, rowcap)
                            genCount += 1
                if self.finData is not None:        # if financial data exists(user already specified and accept the investmentDialog), save the financial data as well
                    self.saveFinData()
                if self.selectedDFSumDiff is not None:
                    self.saveTotalEnergy()
                try:
                    self.writer.save()              # try to save the xlsx file
                    junk = QtWidgets.QMessageBox.information(self.centralwidget, 'Successful',
                                                         'Formatting is saved in ' + path_leaf(file_name)
                                                          + '. File is saved successfully.')
                except:                             # show error if file is opened in another program such as excel
                    junk = QtWidgets.QMessageBox.warning(self.centralwidget, 'Error',
                                                           'Cannot save. File may be opened in another program')
                    self.writer = None
                    self.workbook = None
                    self.formatpercent = None
                    self.format1 = None             # clear all variables relevant to processing the saving


    def insertChart(self, file1, file2, attr, catname, count, df_temp, df_length, rowcap):      # a helper function to save the chart
        strrow = (count + 1) * (df_temp.shape[0] + 3) + rowcap      # row placement of table
        strcol = df_length + 4                                      # column placement of table (to prevent clash of multiple table or charts
        df_temp.to_excel(self.writer, sheet_name=catname, startrow=strrow, startcol=strcol)     # insert the data to xlsx
        file1value = df_temp.loc[file1][len(df_temp.columns) - 1]       # find the value needed for calculation of difference
        file2value = df_temp.loc[file2][len(df_temp.columns) - 1]
        index = df_temp.index.get_loc(file2) + 1        # get the position for writing the difference in xlsx
        self.writer.sheets[catname].write(strrow + index, strcol + len(df_temp.columns) + 1,
                                      (file2value - file1value) / file1value, self.formatpercent)   # write the difference in xlsx with percentage format
        worksheet = self.writer.sheets[catname]
        worksheet.set_column(strcol + 1, strcol + len(df_temp.columns), None, self.format1)     # set the column on the table with numeric format
        chart1 = self.workbook.add_chart({'type': 'column'})            # add a column chart based on the data
        chart1.add_series({'values': [catname, strrow + 1, strcol + len(df_temp.columns),
                                      strrow + 2, strcol + len(df_temp.columns)],       # variables of values and categories are sheetname, startrow, startcolumn
                           'categories': [catname, strrow + 1, strcol, strrow + 2, strcol], # endrow, endcolumn on the xlsx
                           'data_labels': {'value': True},
                           'name': attr + '(' + catname + ')'})
        chart1.set_y_axis({'min': 0, 'name': 'kWh'})
        chart1.set_legend({'position': 'none'})
        worksheet.insert_chart(strrow, strcol + len(df_temp.columns) + 2, chart1)   # insert the chart to the xlsx file

    def saveSavingData(self, df, catname, strcol):           # helper function to save each technologies' energy saving data comparison between different buildings
        df = df.assign(diff=0.00)
        for i in df.index.values:
            building, tech = i.split(' - ')
            if 'Baseline' not in i:
                df.at[i, 'diff'] = df.at[building + ' - Baseline', 'Sum'] - df.at[
                    i, 'Sum']  # calculate the saving from the baseline
        for i in df.index.values:
            if 'Baseline' in i: df.drop([i], inplace=True)  # drop the baseline row
        techProcessed = list()
        rowProcess = 0
        worksheet = self.writer.sheets[catname]
        worksheet.set_column(strcol+1, strcol+len(df.columns), None, self.format1)
        for name in df.index.values:  # add a comparison chart for every technology saving between buildings
            build, tech = name.split(' - ')
            if tech not in techProcessed:
                techProcessed.append(tech)
                data = df.filter(like=tech, axis=0)
                data = data.loc[:, ['diff']]
                if len(data.index) > 1:
                    data.to_excel(self.writer, sheet_name=catname, startrow=rowProcess, startcol = strcol)
                    chart1 = self.workbook.add_chart({'type': 'column'})
                    chart1.add_series(
                        {'values': [catname, rowProcess + 1, strcol + 1,  # values of chart received from payback table
                                    rowProcess + data.shape[0], strcol + 1],
                         'categories': [catname, rowProcess + 1, strcol, rowProcess + data.shape[0], strcol],
                         'data_labels': {'value': True},
                         'name': tech + '(Saving Comparison)'})
                    chart1.set_legend({'position': 'none'})
                    chart1.set_y_axis({'min': 0, 'name': 'Saving %' + tech})
                    worksheet.insert_chart(rowProcess, strcol+len(data.columns) + 2, chart1)
                    rowProcess += data.shape[0] + 1


    def saveFinData(self):              # helper function to save financial data as a separate sheet in the same xlsx file
        self.finData.to_excel(self.writer, sheet_name='Financial')        #output financial data to a new 'Payback' sheet
        yearlist = [cat for cat in self.finData.columns.values.tolist() if 'year' in cat]       #list all columns value with 'year' as substring
        worksheet = self.writer.sheets['Financial']
        worksheet.set_column(1, len(self.finData.columns), None, self.format1)
        strrow = self.finData.shape[0]+3                    #find rowCount of dataframe
        yearData = self.finData.loc[:, yearlist]            #select only payback data to be shown in chart
        yearData = yearData.loc[(yearData!= 0).any(axis=1)]        #removes all rows containing all zeroes before output to excel
        yearData.to_excel(self.writer, sheet_name='Financial', startrow=strrow)   #output payback data to a separate table in the same excel sheet
        countTech = 0                                           #counter for number of valid technology in excel sheet
        for tech in yearData.index.values:                      #loop for every technology (every row)
            countTech += 1
            chart1 = self.workbook.add_chart({'type': 'column'})    #add a new column chart for each technologies
            chart1.add_series({'values': ['Financial', strrow + countTech, 1,                   #values of chart received from payback table
                                          strrow + countTech, len(yearData.columns)],
                               'categories': ['Financial', strrow, 1, strrow, len(yearData.columns) ],
                               'data_labels': {'value': True},
                               'name': tech + '(Payback)'})
            chart1.set_legend({'position': 'none'})
            chart1.set_y_axis({'name': 'Payback'+tech})
            worksheet.insert_chart(strrow + (countTech-1)*(yearData.shape[0]+2), len(yearData.columns) + 2, chart1)
        techProcessed = list()
        rowProcess = strrow + yearData.shape[0]+1
        for name in yearData.index.values:              # add a comparison chart for every technology saving between buildings
            build, tech = name.split(' - ')
            if tech not in techProcessed:
                techProcessed.append(tech)
                data = self.finData.filter(like=tech, axis=0)
                data = data.loc[:, ['diff']]
                if len(data.index) > 1:
                    data.to_excel(self.writer, sheet_name='Financial', startrow=rowProcess)
                    chart1 = self.workbook.add_chart({'type': 'column'})
                    chart1.add_series(
                        {'values': ['Financial', rowProcess+1, 1,  # values of chart received from payback table
                                    rowProcess+data.shape[0], 1],
                         'categories': ['Financial', rowProcess+1, 0, rowProcess+data.shape[0], 0],
                         'data_labels': {'value': True},
                         'name': tech + '(Saving Comparison)'})
                    chart1.set_legend({'position': 'none'})
                    chart1.set_y_axis({'min': 0, 'name': 'Saving %' + tech})
                    worksheet.insert_chart(rowProcess, len(data.columns)+2, chart1)
                    rowProcess += data.shape[0] + 1

    def saveTotalEnergy(self):              # save building energy + chiller energy in separate spreadsheet
        attlist = list()
        for cat in self.clist.cl:
            if cat.name == 'Chiller energy' or cat.name == 'Building energy':
                attlist.extend(cat.members)
        if not attlist: return
        df = pd.DataFrame()
        for file in self.flist:
            availAttList = file.selectedDFSumDiff.columns.intersection(attlist)
            temp = file.selectedDFSumDiff[availAttList]
            df = df.append(temp)
        df.to_excel(self.writer, sheet_name='Total Energy Breakdown')
        self.writer.sheets['Total Energy Breakdown'].set_column(1, len(df.columns), None, self.format1)
        worksheet = self.writer.sheets['Total Energy Breakdown']
        strrow = df.shape[0] + 2
        for file in self.flist:         # generate acmv and energy use breakdown for all buildings in total energy breakdown sheet
            availAttList = file.selectedDFSumDiff.columns.intersection(attlist)
            temp = file.selectedDFSumDiff[availAttList]
            df_acmv = temp.loc[:,[col for col in df.columns if 'ApHVAC' in col]]        # separate the acmv from others
            df_acmv = df_acmv.T.sort_values(file.building_name+' - Baseline', ascending=False).T
            df_acmv_strrow = strrow
            df_acmv.to_excel(self.writer, sheet_name='Total Energy Breakdown', startrow = df_acmv_strrow)   # print the acmv table to excel
            acmvsum = df_acmv.sum(axis='columns')                                                           # calculate the sum of acmv
            df_noacmv = df.loc[:,[col for col in df.columns if 'ApHVAC' not in col]]                        # separate the non-acmv
            df_noacmv = df_noacmv.assign(ACMV = acmvsum)                                                    # assign the acmv sum as new column acmv on noacmv
            df_noacmv = df_noacmv.T.sort_values(file.building_name+' - Baseline', ascending=False).T
            df_noacmv_strrow = df_acmv_strrow + df_acmv.shape[0] + 2
            df_noacmv.to_excel(self.writer, sheet_name='Total Energy Breakdown', startrow= df_noacmv_strrow)  # print energy use table to excel
            chartpie_acmv = self.workbook.add_chart({'type': 'pie'})                                        # create piechart breakdown for acmv
            chartpie_acmv.add_series({'values': ['Total Energy Breakdown', df_acmv_strrow+1,  1, df_acmv_strrow+1,
                                            len(df_acmv.columns)],
                                 'categories': ['Total Energy Breakdown', df_acmv_strrow,  1, df_acmv_strrow,
                                                len(df_acmv.columns)],
                                 'name':  'ACMV breakdown (' + file.building_name + ')',
                                 'data_labels': {'percentage': True}
                                 })
            worksheet.insert_chart(df_acmv_strrow, len(df_acmv.columns) + 1, chartpie_acmv)
            chartpie_noacmv = self.workbook.add_chart({'type': 'pie'})                                      # create piechart breakdown for energy use
            chartpie_noacmv.add_series({'values': ['Total Energy Breakdown', df_noacmv_strrow + 1, 1, df_noacmv_strrow + 1,
                                                 len(df_noacmv.columns)],
                                      'categories': ['Total Energy Breakdown', df_noacmv_strrow, 1, df_noacmv_strrow,
                                                     len(df_noacmv.columns)],
                                      'name': 'Energy Use Breakdown (' + file.building_name + ')',
                                      'data_labels': {'percentage': True}
                                      })
            worksheet.insert_chart(df_noacmv_strrow, len(df_noacmv.columns) + 1, chartpie_noacmv)
            strrow = df_noacmv_strrow + df_noacmv.shape[0] + 2
        count = 0
        for entry in self.chartList:  # process each entry in the chartList obtained from chartDialog
            file1, file2, attr = entry.split(', ')  # find the specification of the chart
            if attr == 'Sum':  # add 'sum' chart for total energy if attribute is sum
                df_temp = df.loc[[file1, file2], :]  # select the data based on the specification  file1, file2
                if len(df_temp.columns):  # check if data exists
                    df_temp = df_temp.T.append(df_temp.loc[:,[col for col in df_temp.columns if 'DX' not in col]].sum(axis='columns').to_frame().T)
                    df_temp.rename({0: 'Sum'}, axis='index', inplace=True)
                    df_temp = df_temp.T  # add the sum data as column
                    df_length = len(df_temp.columns)  # find the length of the table for positioning
                    self.insertChart(file1, file2, attr, 'Total Energy Breakdown', count, df_temp, df_length - 1,
                                    0)  # call insertChart
                    count += 1


    def resetAllWrapper(self):          #reset erases every formatting done by user, so to prevent accidental reset a messagebox is shown
        if self.flist:
            answer = QtWidgets.QMessageBox.question(self.centralwidget, 'Save before reset', 'Any unsaved formatting'
                                                                                     ' will be lost. Proceed to reset'
                                                                                     ' the application?')
            if answer != QtWidgets.QMessageBox.Yes:             #only reset if user answers yes
                return
        self.resetAll()

    def resetAll(self):             #reset every variables to initialization phase, all formatting and file data removed
        self.flist.clear()
        self.clUI.clear()
        self.chartList.clear()
        self.selectedDF = None
        self.index = 0
        self.allAttributes.clear()
        self.selectedAttributes.clear()
        self.tableViewer.setModel(None)
        self.tablePicker.clear()
        self.clist.readCatFile()
        self.load = QtWidgets.QTreeWidgetItem(self.allAttributes)
        self.energy = QtWidgets.QTreeWidgetItem(self.allAttributes)
        self.breakdown = QtWidgets.QTreeWidgetItem(self.allAttributes)
        _translate = QtCore.QCoreApplication.translate
        __sortingEnabled = self.allAttributes.isSortingEnabled()
        self.allAttributes.setSortingEnabled(False)
        self.allAttributes.topLevelItem(0).setText(0, _translate("MainWindow", "General Load"))
        self.allAttributes.topLevelItem(1).setText(0, _translate("MainWindow", "General Energy"))
        self.allAttributes.topLevelItem(2).setText(0, _translate("MainWindow", "Other"))
        self.allAttributes.setSortingEnabled(__sortingEnabled)
        self.addCList()
        self.dialog = None
        self.attList.clear()
        self.writer = None
        self.workbook = None
        self.format1 = None
        self.formatpercent = None
        self.selectedDFSumDiff = None
        self.finData = None


    def addCList(self):                 #add specified categories from category list into the main application
        count = 3
        _translate = QtCore.QCoreApplication.translate
        for cat in self.clist.cl:
            self.clUI.append(QtWidgets.QTreeWidgetItem(self.allAttributes))
            self.allAttributes.topLevelItem(count).setText(0, _translate("MainWindow", cat.name))
            count += 1

    def checkCList(self, itemstr):       #add attributes based on category members from each category in category list
        index = 0
        for cat in self.clist.cl:
            if itemstr in cat.members:
                for n in range(self.clUI[index].childCount()):
                    if self.clUI[index].child(n).text(0) == itemstr:
                        return
                item = QtWidgets.QTreeWidgetItem(self.clUI[index])
                item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsDragEnabled | QtCore.Qt.ItemIsDropEnabled |
                              QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                _translate = QtCore.QCoreApplication.translate
                item.setText(0, _translate('MainWindow', str(itemstr)))
            index += 1

    def openChartDialog(self):      # called to open the chartDialog
        if self.attList:
            if self.dialog.exec_():
                self.chartList.clear()
                for x in range(self.dialog.ui.selectedCharts.count()):          # postprocessing the chartList retrieved from chartDialog if chartDialog is executed
                    self.chartList.append(self.dialog.ui.selectedCharts.item(x).text())

    def openFinDialog(self):        # function to open the financial calculator dialog
        if self.selectedDFSumDiff is not None:      # only open if user has selected attributes
            df = pd.DataFrame()
            baseExist = False
            for file in self.flist:
                if file.baseindex is not None:  # only open if baseline is available (saving is calculated using baseline)
                    baseExist = True
                    df1 = pd.DataFrame()
                    for cat in self.clist.cl:
                        if cat.name == 'Chiller energy' or cat.name == 'Building energy':  # change condition to select different categories for financial calculation
                            df_temp = file.selectedDFSumDiff[file.selectedDFSumDiff.columns.intersection(cat.members)]
                            if len(df_temp.columns):
                                df_temp = df_temp.T.append(df_temp.sum(axis='columns').to_frame().T)
                                df_temp.rename({0: 'Sum'}, axis='index', inplace=True)
                                df_temp = df_temp.T
                                df_temp = df_temp['Sum'].T      # get the sum of the category 'Chiller energy' or 'Building energy' as a one column dataframe
                            if cat.name == 'Chiller energy':
                                df1 = df1.assign(chiller=df_temp)     # assign the data as a new column in a new dataframe with category name as column name
                            else:
                                df1 = df1.assign(building=df_temp)
                    if self.finData is None:
                        df = df.append(df1)
                    else:
                        for index in df1.index:
                            if index not in self.finData.index:
                                df=df.append(df1.loc[[index],:])
            if len(df.columns):     # only executes if the data exists
                df = df.assign(sum=df.sum(axis='columns'))
                data = df.loc[:,['sum']]    # find the sum and get a one column dataframe with only the sum
                data = data.assign(diff = 0.00)
                for i in df.index.values:
                    building, tech = i.split(' - ')
                    if 'Baseline' not in i:
                        data.at[i, 'sum']= data.at[building + ' - Baseline', 'sum']- data.at[i, 'sum']    # calculate the saving from the baseline
                        a = data.at[i, 'sum']
                        b = data.at[building+' - Baseline', 'sum']
                        c = a/b
                        data.at[i, 'diff']= c*100.0 #data.at[i, 'sum'] / data.at[building + ' - Baseline', 'sum']      # calculate the percentage difference in saving
                for i in df.index.values:
                    if 'Baseline' in i: data.drop([i], inplace=True)                                                       # drop the baseline row
                data = data.assign(tariff=0.2156, investCap=0, investOM=0, saving=0)        # add several default financial data for processing in investmentDialog
                #if not self.finLog:         # initialise the investmentDialog if not initialised
                if self.finData is not None:
                    data = pd.concat([self.finData, data])
                self.finLog = QtWidgets.QDialog()
                self.finLog.ui = Ui_InvestmentDialog()
                self.finLog.ui.setupUi(self.finLog, data)   # pass the data to investmentDialog for processing
                if self.finLog.exec_():
                    self.finData = self.finLog.ui.data  # if executed and accepted, the financial data processed is then retrieved from the investmentDialog
            if not baseExist:
                junk = QtWidgets.QMessageBox.warning(self.centralwidget, 'No baseline data',
                                                                        '"Baseline" data not found, unable to calculate '
                                                                        'technology savings.')      # warning about no baseline data

    def openCatDialog(self):                # function to open category editting dialog from mainwindows
        if self.allAttList:
            dict = self.clist.extractDict()
            catLog = QtWidgets.QDialog()
            catLog.ui = Ui_categoryDialog()
            catLog.ui.setupUi(catLog, dict, self.allAttList)
            if catLog.exec_():
                if catLog.ui.reset:         # if user selected reset, then reset category to default
                    self.clist.resetCatFile()
                    self.defaultcat = True
                else:
                    self.clist.saveCatFile(catLog.ui.dict)  # save the editted category data and reset the application to implement the change
                    self.defaultcat = False
                self.resetAll()



    def btnstate(self,b):           # function to handle two conduction gain checkboxes           #btnstate function not used right now (for window and wall cooling load conduction gain)
        if self.flist:
            if b.text() == "Calculate window cooling load conduction gain":     # handle window conduction gain
                if b.isChecked():               # add the attribute if button is checked
                    for building in self.flist:
                        for file in building.fl:
                            if ('Window Cooling Load Conduction Gain (kWh)' not in file.df.columns):        # check if attribute exists in dataframe
                                button = QtWidgets.QMessageBox.warning(self.centralwidget, 'Attribute not found',
                                                                        'Before simulation, check the conduction gain breakdown '
                                                                        'checkbox in output options')
                                b.setCheckState(QtCore.Qt.Unchecked)
                                return #error message
                    parent = self.allAttributes.findItems('Conduction Gain Breakdown', QtCore.Qt.MatchExactly)[0]   # add a new window conduction gain to
                    item = QtWidgets.QTreeWidgetItem(parent)                                                        # conduction gain breakdown category
                    item.setFlags(                                                                                  # in the QTreeWidget if it exists in dataframe
                        QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsDragEnabled | QtCore.Qt.ItemIsDropEnabled |
                        QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                    _translate = QtCore.QCoreApplication.translate
                    item.setText(0, _translate('MainWindow', 'Window Cooling Load Conduction Gain (kWh)'))
                else:                           # remove the attribute if button is unchecked
                    parent = self.allAttributes.findItems('Conduction Gain Breakdown',QtCore.Qt.MatchExactly)[0]
                    child = self.allAttributes.findItems('Window Cooling Load Conduction Gain (kWh)',
                                                                                      QtCore.Qt.MatchExactly|QtCore.Qt.MatchRecursive)[0]
                    index = parent.indexOfChild(child)
                    parent.takeChild(index)         # find the item in QTreeWidget and remove it
                    for x in range(self.selectedAttributes.count()):
                        if self.selectedAttributes.item(x).text() == 'Window Cooling Load Conduction Gain (kWh)':
                            self.selectedAttributes.takeItem(self.selectedAttributes.row(self.selectedAttributes.item(x)))  # remove the item if it exists in selectedAttributes list
                            self.editTable()                # edit the table
                            self.dialog.ui.addEntry(self.flist, self.clist, self.attList)       # edit the chartDialog
                            indexList = list()
                            for index in range(self.dialog.ui.selectedCharts.count()):
                                if 'Window Cooling Load Conduction Gain (kWh)' in self.dialog.ui.selectedCharts.item(index).text():
                                    indexList.append(index)
                            for index in reversed(indexList):
                                self.dialog.ui.selectedCharts.takeItem(index)       # remove any chart specification which contains the window conduction gain
                            self.chartList.clear()
                            for x in range(self.dialog.ui.selectedCharts.count()):
                                self.chartList.append(self.dialog.ui.selectedCharts.item(x).text())     # update the chartList
                            return

            if b.text() == "Calculate wall cooling load conduction gain":       # same thing with the code below, except this one for wall conduction gain
                if b.isChecked():
                    for building in self.flist:
                        for file in building.fl:
                            if ('Wall Cooling Load Conduction Gain (kWh)' not in file.df.columns):
                                button = QtWidgets.QMessageBox.warning(self.centralwidget, 'Attribute not found',
                                                                       'Before simulation, check the conduction gain breakdown '
                                                                       'checkbox in output options')
                                b.setChecked(QtCore.Qt.Unchecked)
                                return  # error message
                    parent = self.allAttributes.findItems('Conduction Gain Breakdown', QtCore.Qt.MatchExactly)[0]
                    item = QtWidgets.QTreeWidgetItem(parent)
                    item.setFlags(
                        QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsDragEnabled | QtCore.Qt.ItemIsDropEnabled |
                        QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                    _translate = QtCore.QCoreApplication.translate
                    item.setText(0, _translate('MainWindow', 'Wall Cooling Load Conduction Gain (kWh)'))
                else:
                    parent = self.allAttributes.findItems('Conduction Gain Breakdown', QtCore.Qt.MatchExactly)[0]
                    parent.takeChild(
                        parent.indexOfChild(self.allAttributes.findItems('Wall Cooling Load Conduction Gain (kWh)',
                                                                         QtCore.Qt.MatchExactly|QtCore.Qt.MatchRecursive)[0]))
                    for x in range(self.selectedAttributes.count()):
                        if self.selectedAttributes.item(x).text() == 'Wall Cooling Load Conduction Gain (kWh)':
                            self.selectedAttributes.takeItem(
                                self.selectedAttributes.row(self.selectedAttributes.item(x)))
                            self.editTable()
                            self.dialog.ui.addEntry(self.flist, self.clist, self.attList)
                            indexList = list()
                            for index in range(self.dialog.ui.selectedCharts.count()):
                                if 'Wall Cooling Load Conduction Gain (kWh)' in self.dialog.ui.selectedCharts.item(
                                        index).text():
                                    indexList.append(index)
                            for index in reversed(indexList):
                                self.dialog.ui.selectedCharts.takeItem(index)
                            self.chartList.clear()
                            for x in range(self.dialog.ui.selectedCharts.count()):
                                self.chartList.append(self.dialog.ui.selectedCharts.item(x).text())
                            return
        else:
            button = QtWidgets.QMessageBox.warning(self.centralwidget, 'File not found, click open to load a file',
                                                   'Before simulation, check the conduction gain breakdown '
                                                   'checkbox in output options')        # warning if no file is opened
            b.setChecked(QtCore.Qt.Unchecked)
            return  # error message

class MyWindow(QtWidgets.QMainWindow):
    def closeEvent(self,event):
        result = QtWidgets.QMessageBox.question(self,
                      "Confirm Exit...",
                      "Are you sure you want to exit ?",
                      QtWidgets.QMessageBox.Yes| QtWidgets.QMessageBox.No)
        event.ignore()

        if result == QtWidgets.QMessageBox.Yes:
            event.accept()


if __name__ == "__main__":          # main execution of the program
    import sys
    app = QtWidgets.QApplication(sys.argv)
    sys.excepthook=excepthook       # create an excepthook to catch any unhandled error
    MainWindow = MyWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)      # setup main window and show it
    MainWindow.show()
    sys.exit(app.exec_())       # exit when app exit
