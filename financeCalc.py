from PyQt5 import QtCore, QtGui, QtWidgets, QtChart
import pandas as pd
import sys
# This module is a side window called by main window class for financial calculation (saving, investment, etc)

class Ui_InvestmentDialog(object):
    def setupUi(self, InvestmentDialog, data):      # setting up widgets in the window (almost like a constructor), takes in a dataframe which has been
        InvestmentDialog.setObjectName("InvestmentDialog")      # preprocessed by main window
        InvestmentDialog.resize(441, 321)
        self.data = data                            # save the dataframe in a class variable

        self.gridLayout = QtWidgets.QGridLayout(InvestmentDialog)
        self.gridLayout.setObjectName("gridLayout")

        self.savingLabel = QtWidgets.QLabel(InvestmentDialog)               # labels are string displayed on interface
        self.savingLabel.setObjectName("savingLabel")                       # display the word 'Saving'
        self.gridLayout.addWidget(self.savingLabel, 3, 0, 2, 1)

        self.savingDollarLabel = QtWidgets.QLabel(InvestmentDialog)
        self.savingDollarLabel.setObjectName("savingDollarLabel")           # display the unit of S$/year
        self.gridLayout.addWidget(self.savingDollarLabel, 4, 2, 1, 1)

        self.electrictyLabel = QtWidgets.QLabel(InvestmentDialog)
        self.electrictyLabel.setObjectName("ElectrictyLabel")               # display the word 'Electricity tariff'
        self.gridLayout.addWidget(self.electrictyLabel, 2, 0, 1, 1)

        self.savingWattLabel = QtWidgets.QLabel(InvestmentDialog)
        self.savingWattLabel.setObjectName("savingWattLabel")               # display the unit of kWh/year
        self.gridLayout.addWidget(self.savingWattLabel, 3, 2, 1, 1)

        self.electricity = QtWidgets.QDoubleSpinBox(InvestmentDialog)       # a doublespinbox for user to change electricity tariff
        self.electricity.setMinimum(0.0001)
        self.electricity.setDecimals(4)                                     # 4 decimal points
        self.electricity.setGroupSeparatorShown(True)                       # add coma as thousand-separator
        self.electricity.setSingleStep(0.001)                               # set changing the value by 0.001 by clicking the arrow on the box
        self.electricity.setProperty("value", 0.2156)
        self.electricity.setObjectName("Electricity")
        self.gridLayout.addWidget(self.electricity, 2, 1, 1, 2)
        self.electricity.valueChanged.connect(self.inputChangeWrapper)      # connect to inputChangeWrapper function for processing input when value is changed

        self.investCapBox = QtWidgets.QDoubleSpinBox(InvestmentDialog)      # a doublespinbox for user to change cost of investment capital
        self.investCapBox.setMaximum(sys.float_info.max)
        self.investCapBox.setGroupSeparatorShown(True)
        self.investCapBox.setObjectName("InvestCapBox")
        self.gridLayout.addWidget(self.investCapBox, 5, 1, 1, 2)
        self.investCapBox.valueChanged.connect(self.capBoxHandler)          # connect to capBoxHandler function when value is changed

        self.investOMBox = QtWidgets.QDoubleSpinBox(InvestmentDialog)       # a doublespinbox for user to change cost of operations and maintenance
        self.investOMBox.setMaximum(sys.float_info.max)
        self.investOMBox.setGroupSeparatorShown(True)
        self.investOMBox.setObjectName("InvestOMBox")
        self.gridLayout.addWidget(self.investOMBox, 7, 1, 1, 2)
        self.investOMBox.valueChanged.connect(self.omBoxHandler)            # connect to omBoxHandler function when value is changed

        self.investCapLabel = QtWidgets.QLabel(InvestmentDialog)
        self.investCapLabel.setObjectName("investCapLabel")                 # label for capital investment
        self.gridLayout.addWidget(self.investCapLabel, 5, 0, 1, 1)

        self.investOMLabel = QtWidgets.QLabel(InvestmentDialog)
        self.investOMLabel.setObjectName("investOMLabel")                   # label for operations and maintenance investment
        self.gridLayout.addWidget(self.investOMLabel, 7, 0, 1, 1)

        self.chartView = QtChart.QChartView(InvestmentDialog)               # chartView to display payback from year 1 to 10
        self.chartView.setObjectName("chartView")
        self.chartView.setMinimumSize(700, 100)
        self.gridLayout.addWidget(self.chartView, 1, 3, 8, 2)

        self.Tech = QtWidgets.QComboBox(InvestmentDialog)                   # a drop down list of all tech in the dataframe (except baseline, as the saving is
        self.Tech.setObjectName("Tech")                                     # compared against baseline)
        self.gridLayout.addWidget(self.Tech, 1, 0, 1, 3)
        for keys in self.data.index.values:                                 # add the tech into the list
            if 'Baseline' not in keys:
                self.Tech.addItem(keys)
        self.Tech.activated.connect(self.techChange)                        # connect to techChange when user selects an item from it,
                                                                            # regardless of whether the item selected change
        self.savingDollar = QtWidgets.QLineEdit(InvestmentDialog)
        self.savingDollar.setReadOnly(True)
        self.savingDollar.setObjectName("SavingDollar")
        self.gridLayout.addWidget(self.savingDollar, 4, 1, 1, 1)            # show value of the saving in dollar (cannot be changed by user)

        self.savingWatt = QtWidgets.QLineEdit(InvestmentDialog)
        self.savingWatt.setReadOnly(True)
        self.savingWatt.setObjectName("SavingWatt")
        self.gridLayout.addWidget(self.savingWatt, 3, 1, 1, 1)              # show value of saving in kWh taken from dataframe (cannot be changed by user)

        self.investmentOMSlider = QtWidgets.QSlider(InvestmentDialog)       # create a horizontal slider which is able to change value of Investment(ops&m)
        self.investmentOMSlider.setMaximum(20000)
        self.investmentOMSlider.setOrientation(QtCore.Qt.Horizontal)
        self.investmentOMSlider.setObjectName("InvestmentOMSlider")
        self.gridLayout.addWidget(self.investmentOMSlider, 8, 0, 1, 3)
        self.investmentOMSlider.sliderMoved.connect(self.omSliderHandler)   # connect the signal sliderMoved to the function omSliderHandler

        self.investmentCapSlider = QtWidgets.QSlider(InvestmentDialog)      # create a horizontal slider which is able to change value of Investment(Capital)
        self.investmentCapSlider.setMaximum(100000)
        self.investmentCapSlider.setOrientation(QtCore.Qt.Horizontal)
        self.investmentCapSlider.setObjectName("InvestmentCapSlider")
        self.gridLayout.addWidget(self.investmentCapSlider, 6, 0, 1, 3)
        self.investmentCapSlider.sliderMoved.connect(self.capSliderHandler)

        self.buttonBox = QtWidgets.QDialogButtonBox(InvestmentDialog)       # create a set of default buttons in the interface
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Save)    # only cancel and save button are created
        self.buttonBox.setObjectName("buttonBox")
        self.gridLayout.addWidget(self.buttonBox, 9, 2, 1, 3)

        self.retranslateUi(InvestmentDialog)
        self.buttonBox.accepted.connect(InvestmentDialog.accept)        # connect save button to dialog accept
        self.buttonBox.rejected.connect(InvestmentDialog.reject)        # connect cancel button to dialog reject
        QtCore.QMetaObject.connectSlotsByName(InvestmentDialog)

    def retranslateUi(self, InvestmentDialog):
        _translate = QtCore.QCoreApplication.translate
        InvestmentDialog.setWindowTitle(_translate("InvestmentDialog", "Financial Calculator"))
        self.savingLabel.setText(_translate("InvestmentDialog", "Saving"))
        self.savingDollarLabel.setText(_translate("InvestmentDialog", "S$/year"))
        self.electrictyLabel.setText(_translate("InvestmentDialog", "Electricity tariff"))
        self.savingWattLabel.setText(_translate("InvestmentDialog", "kWh/year"))
        self.electricity.setSuffix(_translate("InvestmentDialog", "S$/kWh"))
        self.investCapBox.setSuffix(_translate("InvestmentDialog", "S$"))
        self.investOMBox.setSuffix(_translate("InvestmentDialog", "S$"))
        self.investCapLabel.setText(_translate("InvestmentDialog", "Investment (capital)"))
        self.investOMLabel.setText(_translate("InvestmentDialog", "Investment (Ops&M)"))

    def capBoxHandler(self, double):                        # function to handle capBox value change (capSlider value is also changed)
        self.investmentCapSlider.blockSignals(True)
        self.investmentCapSlider.setValue(double)
        self.investmentCapSlider.blockSignals(False)
        self.inputChangeWrapper()

    def omBoxHandler(self, double):                         # function to handle omBox value change (omSlider value is also changed)
        self.investmentOMSlider.blockSignals(True)
        self.investmentOMSlider.setValue(double)
        self.investmentOMSlider.blockSignals(False)
        self.inputChangeWrapper()

    def capSliderHandler(self, number):                     # function to handle capSlider moved (value in capBox is changed accordingly)
        self.investCapBox.blockSignals(True)
        self.investCapBox.setValue(number)
        self.investCapBox.blockSignals(False)
        self.inputChangeWrapper()

    def omSliderHandler(self, number):                      # function to handle omSlider moved (value in omBox is changed accordingly)
        self.investOMBox.blockSignals(True)
        self.investOMBox.setValue(number)
        self.investOMBox.blockSignals(False)
        self.inputChangeWrapper()

    def inputChangeWrapper(self):                           # function to handle all change of inputs except technology drop down list
        tech = self.Tech.currentText()
        saveWatt = self.data.at[tech, 'sum']
        self.savingWatt.setText("{0:7,.2f}".format(saveWatt))   # savingWatt text shown is changed
        elecValue = self.electricity.value()
        self.data.at[tech, 'tariff'] = elecValue                # tariff in the dataframe is changed according to electricity value from interface
        investCap = self.investCapBox.value()
        self.data.at[tech, 'investCap'] = investCap             # same as tariff, but for capital investment
        investOM = self.investOMBox.value()
        self.data.at[tech, 'investOM'] = investOM               # same as tariff, but for operations and maintenance investment
        saveDollar = self.savingDollarHandler(saveWatt, elecValue)          # calculate saveDollar by calling savingDollarHandler
        self.data.at[tech,'saving'] = saveDollar                # change data of saving in the dataframe
        self.calculatePaybackROI()                              # calculate year 1 - 10 payback and roi

    def techChange(self):                                       # function to handle drop down list change technology, set value from dataframe to value displays
        tech = self.Tech.currentText()                          # in the interface
        saveWatt = self.data.at[tech, 'sum']
        self.savingWatt.setText("{0:7,.2f}".format(saveWatt))
        elecValue = self.data.at[tech, 'tariff']
        self.electricity.blockSignals(True)
        self.electricity.setValue(elecValue)
        self.electricity.blockSignals(False)
        self.investCapBox.blockSignals(True)
        self.investCapBox.setValue(self.data.at[tech, 'investCap'])
        self.investCapBox.blockSignals(False)
        self.investOMBox.blockSignals(True)
        self.investOMBox.setValue(self.data.at[tech, 'investOM'])
        self.investOMBox.blockSignals(False)
        self.investmentCapSlider.blockSignals(True)
        self.investmentCapSlider.setValue(self.data.at[tech, 'investCap'])
        self.investmentCapSlider.blockSignals(False)
        self.investmentOMSlider.blockSignals(True)
        self.investmentOMSlider.setValue(self.data.at[tech, 'investOM'])
        self.investmentOMSlider.blockSignals(False)
        saveDollar = self.savingDollarHandler(saveWatt, elecValue)
        self.createChart()

    def savingDollarHandler(self, saveWatt, elecValue):             # handle calculation and change in saveDollar
        saveDollar = saveWatt * elecValue
        self.savingDollar.setText("{0:7,.2f}".format(saveDollar))
        return saveDollar

    def calculatePaybackROI(self):          # handles the calculation and generation of year 1 - 10 value of payback and roi appended to dataframe
        payback = list()
        roi = list()
        for year in range(1, 11):
            for tech in self.data.index.values:
                investment = self.data.at[tech,'investCap']+year*self.data.at[tech,'investOM']  #calculate yearly investment from data in dataframe
                if investment == 0:                         # if investment is 0, data is unavailable, so paybackvalue and roivalue are set to 0
                    paybackvalue = 0
                    roivalue = 0
                else:                                       # else set them properly (OpsnM investment are multiplied with year)
                    paybackvalue = self.data.at[tech,'saving']*year - (self.data.at[tech,'investCap']+year*self.data.at[tech,'investOM'])
                    roivalue = paybackvalue/(self.data.at[tech,'investCap']+year*self.data.at[tech,'investOM'])
                payback.append(paybackvalue)
                roi.append(roivalue)
            if len(self.data.columns.intersection(['year'+str(year), 'roi'+str(year)])):        # if the column exists in dataframe, then change the value
                self.data.loc[:, 'year'+str(year)]=payback
                self.data.loc[:, 'roi'+str(year)]=roi
            else:                                                                   # if it does not exists, then create the column using assign
                self.data = self.data.assign(payyear=payback)
                self.data = self.data.assign(roiyear = roi)
                yearname = 'year'+str(year)
                roiname = 'roi'+str(year)
                self.data.rename(index = str, columns = {'payyear': yearname, 'roiyear': roiname}, inplace=True)
            payback.clear()
            roi.clear()
        self.createChart()              # call create chart to create and show the chart in interface

    def createChart(self):              # function to create a barchart similar to a time series, with each bar showing each year 1 - 10 payback value
        tech = self.Tech.currentText()
        series = QtChart.QBarSeries()
        chart = QtChart.QChart()
        axis = QtChart.QBarCategoryAxis(chart)
        categories = [cat for cat in self.data.columns.values.tolist() if 'year' in cat]
        set0 = QtChart.QBarSet('Payback')
        list0 = list()
        for cat in categories:
            list0.append(self.data.at[tech, cat])
        set0.append(list0)
        series.append(set0)
        #list0.clear()
        #roicategories = [cat for cat in self.data.columns.values.tolist() if 'roi' in cat]
        #set1 = QtChart.QBarSet('ROI')
        #for cat in roicategories:
            #list0.append(self.data.at[tech, cat])
        #set1.append(list0)
        #series.append(set1)
        series.setLabelsVisible()
        chart.addSeries(series)
        axis.append(categories)
        chart.createDefaultAxes()
        chart.setAxisX(axis, series)
        self.chartView.setChart(chart)



