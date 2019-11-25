from PyQt5 import QtCore, QtGui, QtWidgets, QtChart

# This module is a side window/dialog class for chart selection. It uses QT to generate the interface and is called by main window.

class Ui_Dialog(object):
    def setupUi(self, Dialog, flist, clist, attList):       # setting up the widgets inside the window, receiving file list, category list, and
        self.dialog = Dialog                                # selected attribute list from main window when initialized
        Dialog.setObjectName("Dialog")
        Dialog.resize(600, 400)
        self.fl = flist
        self.cl = clist

        self.gridLayout_2 = QtWidgets.QGridLayout(Dialog)
        self.gridLayout_2.setObjectName("gridLayout_2")

        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")

        self.techselector1 = QtWidgets.QComboBox(Dialog)            # drop down selectable list/combobox to select which technology/set of data is shown
        self.techselector1.setObjectName("techselector1")
        self.techselector1.activated.connect(self.showChart)        # when drop down is activated (changed or not) the chart shown is changed accordingly
        self.gridLayout.addWidget(self.techselector1, 0, 0, 1, 2)

        self.techselector2 = QtWidgets.QComboBox(Dialog)            # same as techselector1, to make a comparison between the two technology
        self.techselector2.setObjectName("techselector2")
        self.techselector2.activated.connect(self.showChart)
        self.gridLayout.addWidget(self.techselector2, 1, 0, 1, 2)

        self.attselector = QtWidgets.QComboBox(Dialog)              # drop down selectable list for selecting which attribute to compare on
        self.attselector.setObjectName("comboBox_3")
        self.attselector.activated.connect(self.showChart)
        self.gridLayout.addWidget(self.attselector, 2, 0, 1, 2)

        for file in flist:                                          # adding items to the drop down lists according to file list and attribute list
            for sheet in file.fl:
                self.techselector1.addItem(sheet.sheet_name)                       # temporary change
                self.techselector2.addItem(sheet.sheet_name)
        if attList:
            self.attselector.addItem('Sum')                         # adding a special item to attribute to let user compare total of each category
        for att in attList:
            self.attselector.addItem(att)

        self.chartView = QtChart.QChartView(Dialog)                 # chartView to show chart on window (responsive and simultaneous)
        self.chartView.setObjectName("chartView")
        self.gridLayout.addWidget(self.chartView, 3, 0, 6, 1)

        self.selectedCharts = QtWidgets.QListWidget(Dialog)         # maintain a list of user selected chart specification for user to see, save, add, and remove
        self.selectedCharts.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.selectedCharts.setObjectName("selectedCharts")
        for file in flist:
            if file.baseline is not None:                            # only adds comparison between baseline and other technologies on the sum if baseline data
                for sheet in file.fl:                                 # exists. Comparison chart for these are added automatically by program, but user can still
                    if 'Baseline' not in sheet.sheet_name:                   # delete them
                        item = QtWidgets.QListWidgetItem()
                        item.setFlags(
                            QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                        self.selectedCharts.addItem(item)
                        _translate = QtCore.QCoreApplication.translate
                        item.setText(_translate("Dialog", file.building_name + ' - ' + 'Baseline' + ", " + sheet.sheet_name + ", " + 'Sum'))     # text in the list is set 'tech1, tech2, att'

        self.gridLayout.addWidget(self.selectedCharts, 3, 1, 6, 1)

        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Save     # four buttons exists 'Save, Discard, Cancel, Apply'
                                          |QtWidgets.QDialogButtonBox.Discard|QtWidgets.QDialogButtonBox.Apply)
        self.buttonBox.setObjectName("buttonBox")
        self.gridLayout.addWidget(self.buttonBox, 9, 0, 1, 2)

        self.gridLayout_2.addLayout(self.gridLayout, 0, 0, 1, 1)

        self.retranslateUi(Dialog)
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Apply).clicked.connect(self.saveChart)         # apply the user selected entry from drop down lists as a new entry in selected list
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Cancel).clicked.connect(Dialog.reject)         # closes the window, does not destroy it, all data is saved but not added to main window
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Discard).clicked.connect(self.discardChart)    # remove an entry from selected list selected by the user
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Save).clicked.connect(Dialog.accept)           # same as cancel, but main window grabs the data from this window for saving to xlsx

        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Charts Selection"))

    def saveChart(self):                                # function to save user-selected tech and att from drop down list to selectedCharts list (add entry)
        if self.attselector.currentText():              # checks if any attribute is selected by user
            file1 = self.techselector1.currentText()
            file2 = self.techselector2.currentText()
            att = self.attselector.currentText()            # takes all the currentTexts from drop down lists to create a new entry
            for x in range(self.selectedCharts.count()):
                if self.selectedCharts.item(x).text() == file1 + ", " + file2 + ", " + att or \
                        self.selectedCharts.item(x).text() == file2 + ", " + file1 + ", " + att:          # checks whether the new entry is already selected in
                    junk = QtWidgets.QMessageBox.warning(self.dialog, 'Chart specification exists',         # the list, if it is, it won't be added
                                                       'The selected chart specification is already saved, '
                                                       'select a different chart specification to save')
                    return
            item = QtWidgets.QListWidgetItem()
            item.setFlags(
                QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            self.selectedCharts.addItem(item)
            _translate = QtCore.QCoreApplication.translate
            item.setText(_translate("Dialog", file1 + ", " + file2 + ", " + att))         # add the new entry as a new QListWidgetItem

    def discardChart(self):                     # function to remove user selected items from the selectedCharts list
        if self.selectedCharts.selectedItems():
            selectedSC = self.selectedCharts.selectedItems()
            for item in selectedSC:
                self.selectedCharts.takeItem(self.selectedCharts.row(item))

    def addEntry(self, flist, clist, attlist):      # function to add new entries to the drop down list when window is already initialized but called again
        self.fl = flist                             # from main window
        self.cl = clist
        self.techselector1.clear()
        self.techselector2.clear()
        self.attselector.clear()                    # all the drop down lists are first cleared for easier insertion of items
        for file in flist:
            for sheet in file.fl:
                self.techselector1.addItem(sheet.sheet_name)                       # temporary change
                self.techselector2.addItem(sheet.sheet_name)
        if attlist:
            self.attselector.addItem('Sum')
        for att in attlist:
            self.attselector.addItem(att)

    def showChart(self):                            # function to show the chart of data of user selected tech and attribute on QChartView
        file1 = self.techselector1.currentText()
        current_building1, current_tech1 = file1.split(' - ')
        current_buildingI1, current_buildingI2, current_techI1, current_techI2 = 0, 0, 0, 0
        for index, file in enumerate(self.fl):
            if file.building_name == current_building1:
                current_buildingI1 = index
        file2 = self.techselector2.currentText()
        current_building2, current_tech2 = file2.split(' - ')
        for index, file in enumerate(self.fl):
            if file.building_name == current_building2:
                current_buildingI2 = index
        att = self.attselector.currentText()        # assign user selected texts from drop down lists to parameters
        series = QtChart.QBarSeries()               # the lines below creates the bar chart with the 2 tech selected as the categories (x-axis) with data from
        chart = QtChart.QChart()                    # filelist dfSumDiff (summary of sum of each attributes in all dataframes)
        axis = QtChart.QBarCategoryAxis(chart)
        categories = [file1, file2]
        if att != 'Sum':
            set0 = QtChart.QBarSet(att)             # use only the user-specified attribute sum value attained from dataframe as one set of bars
            set0.append([self.fl[current_buildingI1].dfSumDiff.at[file1, att], self.fl[current_buildingI2].dfSumDiff.at[file2, att]])
            series.append(set0)
        else:
            for cat in self.cl.cl:                  # add different set of bars for different category total
                set0 = QtChart.QBarSet(att+'('+cat.name+')')
                df_temp = self.fl[current_buildingI1].dfSumDiff[self.fl[current_buildingI1].dfSumDiff.columns.intersection(cat.members)]
                if current_buildingI1 != current_buildingI2:
                    df_temp = df_temp.append(self.fl[current_buildingI2].dfSumDiff[self.fl[current_buildingI2].dfSumDiff.columns.intersection(cat.members)])
                if cat.name != 'Cooling load':
                    df_temp = df_temp.T.append(df_temp.loc[:, [col for col in df_temp.columns if 'DX' not in col]].sum(
                        axis='columns').to_frame().T)  # add the sum of the category as a new column
                else:
                    df_temp = df_temp.T.append(
                        df_temp.loc[:, ['Ap Sys chillers load (kWh)', 'ApHVAC chillers load (kWh)']].sum(
                            axis='columns').to_frame().T)
                df_temp.rename({0: 'Sum'}, axis='index', inplace=True)
                df_temp = df_temp.T
                set0.append([df_temp.at[file1, 'Sum'], df_temp.at[file2, 'Sum']])
                series.append(set0)
        series.setLabelsVisible()
        chart.addSeries(series)
        axis.append(categories)
        chart.createDefaultAxes()
        chart.setAxisX(axis, series)
        self.chartView.setChart(chart)





