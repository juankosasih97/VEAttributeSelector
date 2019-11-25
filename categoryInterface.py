# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'categoryEdit.ui'
#
# Created by: PyQt5 UI code generator 5.10
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import json
from path_leaf import path_leaf

class Ui_categoryDialog(object):
    def setupUi(self, categoryDialog, dict,attList):
        categoryDialog.setObjectName("categoryDialog")
        categoryDialog.resize(458, 601)
        self.dialog = categoryDialog
        self.dict = dict
        self.gridLayout = QtWidgets.QGridLayout(categoryDialog)
        self.gridLayout.setObjectName("gridLayout")
        self.currentCat = None
        self.reset = False

        self.categoryPicker = QtWidgets.QComboBox(categoryDialog)       # qcombobox to select which category to be editted
        self.categoryPicker.setObjectName("categoryPicker")
        self.categoryPicker.setEditable(True)
        self.gridLayout.addWidget(self.categoryPicker, 0, 0, 1, 3)
        for key in dict.keys():                                         # add each existing category to the qcombobox
            self.categoryPicker.addItem(key)
        self.currentCat = self.categoryPicker.currentText()
        self.categoryPicker.editTextChanged.connect(self.renameCat)
        self.categoryPicker.currentIndexChanged.connect(self.changeIndex)

        #self.categoryNameEdit = QtWidgets.QLineEdit(categoryDialog)     # lineedit to edit the name of the selected category
        #self.categoryNameEdit.setObjectName("categoryNameEdit")
        #self.gridLayout.addWidget(self.categoryNameEdit, 2, 0, 1, 3)

        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.horizontalLayout.setObjectName("horizontalLayout")

        self.newButton = QtWidgets.QPushButton(categoryDialog)          # 'new' button to create a new category
        self.newButton.setObjectName("newButton")
        self.newButton.clicked.connect(self.createNewCat)
        self.horizontalLayout.addWidget(self.newButton)

        #self.renameButton = QtWidgets.QPushButton(categoryDialog)       # 'rename' button to rename selected category
        #self.renameButton.setObjectName("renameButton")
        #self.horizontalLayout.addWidget(self.renameButton)

        self.removeCatButton = QtWidgets.QPushButton(categoryDialog)    # 'remove category' button to remove selected category
        self.removeCatButton.setObjectName("removeCatButton")
        self.removeCatButton.clicked.connect(self.removeCat)
        self.horizontalLayout.addWidget(self.removeCatButton)

        self.gridLayout.addLayout(self.horizontalLayout, 2, 0, 1, 3)

        self.allAttrList = QtWidgets.QListWidget(categoryDialog)        # a listwidget to display all available attributes
        self.allAttrList.setObjectName("allAttrList")
        self.allAttrList.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.allAttrList.itemDoubleClicked.connect(self.addAttDC)
        for att in sorted(attList):
            item = QtWidgets.QListWidgetItem()
            item.setFlags(
                QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            self.allAttrList.addItem(item)
            item.setText(att)
        self.gridLayout.addWidget(self.allAttrList, 3, 0, 6, 1)

        self.catAttrList = QtWidgets.QListWidget(categoryDialog)        # a listwidget to display all attributes under selected category
        self.catAttrList.setObjectName("catAttrList")
        self.catAttrList.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        for att in dict[self.currentCat]:
            item = QtWidgets.QListWidgetItem()
            item.setFlags(
                QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            self.catAttrList.addItem(item)
            item.setText(att)
        self.gridLayout.addWidget(self.catAttrList, 3, 2, 6, 1)

        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")

        self.openCatButton = QtWidgets.QPushButton(categoryDialog)          # 'open' button to call filedialog to parse a category configuration file
        self.openCatButton.setObjectName("openCatButton")
        self.openCatButton.clicked.connect(self.openCatFile)
        self.verticalLayout.addWidget(self.openCatButton)

        self.addButton = QtWidgets.QPushButton(categoryDialog)          # 'add' button to add selected attributes from allAttrList to catAttrList
        self.addButton.setObjectName("addButton")
        self.addButton.clicked.connect(self.addAtt)
        self.verticalLayout.addWidget(self.addButton)

        self.removeButton = QtWidgets.QPushButton(categoryDialog)       # 'remove' button to remove selected attributes from catAttrList
        self.removeButton.setObjectName("removeButton")
        self.removeButton.clicked.connect(self.removeAtt)
        self.verticalLayout.addWidget(self.removeButton)

        self.gridLayout.addLayout(self.verticalLayout, 4, 1, 5, 1)

        self.buttonBox = QtWidgets.QDialogButtonBox(categoryDialog)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok|QtWidgets.QDialogButtonBox.Reset|QtWidgets.QDialogButtonBox.Save)  # Ok and cancel button for exiting dialog
        self.buttonBox.setObjectName("buttonBox")
        self.gridLayout.addWidget(self.buttonBox, 10, 0, 1, 3)
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Reset).clicked.connect(self.resetAllWrapper)


        self.retranslateUi(categoryDialog)
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Ok).clicked.connect(self.acceptWrapper)
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Cancel).clicked.connect(self.cancelWrapper)
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Save).clicked.connect(self.saveCatFile)
        QtCore.QMetaObject.connectSlotsByName(categoryDialog)

    def retranslateUi(self, categoryDialog):
        _translate = QtCore.QCoreApplication.translate
        categoryDialog.setWindowTitle(_translate("categoryDialog", "Dialog"))
        self.newButton.setText(_translate("categoryDialog", "Add New Category"))
        #self.renameButton.setText(_translate("categoryDialog", "Rename Category"))
        self.openCatButton.setText(_translate("categoryDialog", "Open Category File"))
        self.removeCatButton.setText(_translate("categoryDialog", "Remove Category"))
        self.addButton.setText(_translate("categoryDialog", "Add Attr"))
        self.removeButton.setText(_translate("categoryDialog", "Remove Attr"))

    def createNewCat(self):                 # create a new empty category and select it
        self.categoryPicker.addItem('new category')
        self.dict['new category']=list()
        self.categoryPicker.setCurrentIndex(self.categoryPicker.findText('new category', QtCore.Qt.MatchExactly))

    def changeIndex(self, ind):             # activated when index of qcombobox is changed, either by user or program,
        if ind != -1:
            self.currentCat = self.categoryPicker.itemText(ind)     # change the content of catAttrList according to dict.
            self.catAttrList.clear()
            for att in self.dict[self.currentCat]:
                item = QtWidgets.QListWidgetItem()
                item.setFlags(
                    QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                self.catAttrList.addItem(item)
                item.setText(att)
        else:
            self.currentCat = None
            self.catAttrList.clear()

    def renameCat(self, string):            # rename currently selected category using lineedit and implement the change in dict
        self.categoryPicker.blockSignals(True)      # signals are blocked during processing of changes
        if string not in self.dict.keys():
            if self.currentCat is not None:
                self.dict[string] = self.dict.pop(self.currentCat)      # change the key in dict to string from currentCat
                self.categoryPicker.setItemText(self.categoryPicker.currentIndex(), string)     # set the item text as string in drop down list
                self.currentCat = string            # set the currentCat as string
        self.categoryPicker.blockSignals(False)


    def removeCat(self):    # remove currently selected category from qcombobox and delete its data from dict
        if self.currentCat is not None:
            answer = QtWidgets.QMessageBox.question(self.dialog, 'Remove category?', 'Are you sure you want to remove the selected categories?'
                                                                                             ' Selected category data will be lost.')
            if answer != QtWidgets.QMessageBox.Yes:  # only remove if user answers yes
                return
            del self.dict[self.categoryPicker.currentText()]
            self.categoryPicker.removeItem(self.categoryPicker.currentIndex())

    def addAttDC(self, item):       # add item to catAttrList by doubleclicking on item in allAttrList
        stritem = item.text()
        for x in range(self.catAttrList.count()):
            if self.catAttrList.item(x).text() == stritem:  # if item is in catAttrList, then don't add anymore
                return
        item = QtWidgets.QListWidgetItem()
        item.setFlags(
            QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
        self.catAttrList.addItem(item)  # else add the item to catAttrList
        item.setText(stritem)
        self.dict[self.currentCat].append(stritem)  # add the item under the current category in dict


    def addAtt(self):       # add selected items from allAttrList to catAttrList, implement the change on the specific
        if self.currentCat is not None and self.allAttrList.selectedItems():
            for att in self.allAttrList.selectedItems():    # key-value pair in the dict
                self.addAttDC(att)

    def removeAtt(self):
        if self.currentCat is not None and self.catAttrList.selectedItems():
            for att in self.catAttrList.selectedItems():
                self.catAttrList.takeItem(self.catAttrList.row(att))
                self.dict[self.currentCat].remove(att.text())

    def acceptWrapper(self):
        answer = QtWidgets.QMessageBox.question(self.dialog, 'Reset application?',
                                                'In order to arrange the change in categories, the application will need'
                                                ' to reset. Saving will also disable any formatting feature'
                                                ' exclusive for default category setting. Any unsaved formatting data will be lost. Reset anyway?')
        if answer != QtWidgets.QMessageBox.Yes:  # only remove if user answers yes
            return
        self.dialog.accept()

    def cancelWrapper(self):
        answer = QtWidgets.QMessageBox.question(self.dialog, 'Cancel category editting?',
                                                'Canceling will remove any editting of categories. Are you sure?')
        if answer != QtWidgets.QMessageBox.Yes:  # only remove if user answers yes
            return
        self.dialog.reject()

    def resetAllWrapper(self):
        answer = QtWidgets.QMessageBox.question(self.dialog, 'Reset category to default?',
                                                'Are you sure you want to reset category to default? This window will close to implement the change.')
        if answer != QtWidgets.QMessageBox.Yes:  # only remove if user answers yes
            return
        self.reset = True
        self.dialog.accept()

    def openCatFile(self):
        fdlg = QtWidgets.QFileDialog()
        fdlg.setFileMode(QtWidgets.QFileDialog.ExistingFile)  # only enable user to select one existing file
        fdlg.setNameFilter('Text file(*.txt)')  # only enable to open xlsx file
        if fdlg.exec_():
            openfile = None
            data = None
            try:
                openfile = fdlg.selectedFiles()[0]
                with open(openfile, "rt") as fp:
                    data = json.load(fp)
            except:
                warn = QtWidgets.QMessageBox.warning(self.dialog, 'Text file cannot be read',
                                                               'The selected file ' + path_leaf(openfile) + ' cannot be read as as categories, '
                                                               'select a different category configuration file')
                return
            if data is not None:
                self.dict = data
                self.categoryPicker.blockSignals(True)
                self.catAttrList.blockSignals(True)
                self.categoryPicker.clear()
                self.catAttrList.clear()
                for key in self.dict.keys():  # add each existing category to the qcombobox
                    self.categoryPicker.addItem(key)
                for att in self.dict[self.currentCat]:
                    item = QtWidgets.QListWidgetItem()
                    item.setFlags(
                        QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                    self.catAttrList.addItem(item)
                    item.setText(att)
                self.categoryPicker.blockSignals(False)
                self.catAttrList.blockSignals(False)

    def saveCatFile(self):
        sfdlg = QtWidgets.QFileDialog()
        sfdlg.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)  # call a QFileDialog which saves only txt file
        sfdlg.setNameFilter('Text file(*.txt)')
        if sfdlg.exec_():
            file_name = sfdlg.selectedFiles()[0]
            with open(file_name, "wt") as fp:
                json.dump(self.dict, fp)







    #if __name__ == "__main__":          # main execution of the program
    #import sys
    #dict = {}
    #app = QtWidgets.QApplication(sys.argv)
    #sys.excepthook=excepthook       # create an excepthook to catch any unhandled error
    #MainWindow = QtWidgets.QDialog()
    #ui = Ui_categoryDialog()
    #ui.setupUi(MainWindow, dict)      # setup main window and show it
    #MainWindow.show()
    #sys.exit(app.exec_())       # exit when app exit
