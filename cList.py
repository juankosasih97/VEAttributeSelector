from category import Category
import json

# This module cList.py contains class cList (category list)

class cList:
    def __init__(self):                 # constructor (all specific categories in the program are defined here
        self.cl = list()
        #self.catFileReadable = False
        #try:
        #    self.readCatFile()
        #    self.catFileReadable = True
        #except IOError:
        self.addCat('Cooling load')
        self.addMembertoCat('Cooling load', 'Chillers load (kWh)')                              # remove from sum
        self.addMembertoCat('Cooling load', 'Ap Sys chillers load (kWh)')
        self.addMembertoCat('Cooling load', 'ApHVAC chillers load (kWh)')
        self.addMembertoCat('Cooling load', 'ApHVAC DX cooling systems load (kWh)')             # remove from sum
        self.addCat('Chiller energy')
        self.addMembertoCat('Chiller energy', 'ApHVAC chillers energy (kWh)')
        self.addMembertoCat('Chiller energy', 'ApHVAC DX cooling systems energy (kWh)')            #removed
        self.addMembertoCat('Chiller energy', 'ApHVAC distr pumps energy (kWh)')
        self.addMembertoCat('Chiller energy', 'ApHVAC heat rej fans/pumps energy (kWh)')
        self.addCat('Building energy')
            #self.addMembertoCat('Building energy', 'ApHVAC DX cooling systems energy (kWh)')        # remove from sum
        self.addMembertoCat('Building energy', 'ApHVAC distr fans energy (kWh)')
        self.addMembertoCat('Building energy', 'Lights Misc. A (kWh)')
        self.addMembertoCat('Building energy', 'Equip Misc. H (kWh)')
        self.addMembertoCat('Building energy', 'Other Process (kWh)')
        self.addMembertoCat('Building energy', 'MEC elevators energy (kWh)')
            # self.addCat('Conduction Gain Breakdown')
            # self.addMembertoCat('Conduction Gain Breakdown', 'Conduction gain - external walls (kWh)')#add more categories here
            # self.addMembertoCat('Conduction Gain Breakdown', 'Solar gain (kWh)')
            # self.addMembertoCat('Conduction Gain Breakdown', 'Conduction gain - external windows (kWh)')
            # self.addMembertoCat('Conduction Gain Breakdown', 'ApHVAC chillers load (kWh)')
            # self.addMembertoCat('Conduction Gain Breakdown', 'Window Cooling Load Conduction Gain (kWh)')
            # self.addMembertoCat('Conduction Gain Breakdown', 'Wall Cooling Load Conduction Gain (kWh)')#feature maybe tba for adding categories on UI
        self.saveCatFile()


    def addCat(self, name):             # function to append a new category to the list
        self.cl.append(Category(name))

    def addMembertoCat(self, name, membername):         # function to append a string 'membername' as member of category with name 'name'
        for x in range(len(self.cl)):
            if self.cl[x].name == name :
                self.cl[x].addMembers(membername)
                return

    def getMember(self, name):
        for cat in self.cl:
            if cat.name == name:
                return cat.members


    ## section still in progress

    def addDict(self, dict):                    # function to import category and its members from a dictionary
        self.cl.clear()
        for cat in dict:
            self.addCat(cat)
            for att in dict[cat]:
                self.addMembertoCat(cat, att)

    def extractDict(self):                      # function to extract from list of categories in the program to dictionary
        data = {}
        for cat in self.cl:
            data[cat.name] = cat.members
        return data

    def saveCatFile(self, dict = None):         # function to save category data to text file
        if dict is None:
            dict = self.extractDict()
        with open('savecat.txt', "wt") as fp:
            json.dump(dict, fp)

    def readCatFile(self):                      # function to read a category text file to dictionary for importing to program using addDict()
        with open('savecat.txt', "rt") as fp:
            data = json.load(fp)
        self.addDict(data)
        #print("Data: %s" % data)

    def resetCatFile(self):
        self.cl.clear()
        self.addCat('Cooling load')
        self.addMembertoCat('Cooling load', 'Chillers load (kWh)')  # remove from sum
        self.addMembertoCat('Cooling load', 'Ap Sys chillers load (kWh)')
        self.addMembertoCat('Cooling load', 'ApHVAC chillers load (kWh)')
        self.addMembertoCat('Cooling load', 'ApHVAC DX cooling systems load (kWh)')  # remove from sum
        self.addCat('Chiller energy')
        self.addMembertoCat('Chiller energy', 'ApHVAC chillers energy (kWh)')
        self.addMembertoCat('Chiller energy', 'ApHVAC DX cooling systems energy (kWh)')  # removed
        self.addMembertoCat('Chiller energy', 'ApHVAC distr pumps energy (kWh)')
        self.addMembertoCat('Chiller energy', 'ApHVAC heat rej fans/pumps energy (kWh)')
        self.addCat('Building energy')
        # self.addMembertoCat('Building energy', 'ApHVAC DX cooling systems energy (kWh)')        # remove from sum
        self.addMembertoCat('Building energy', 'ApHVAC distr fans energy (kWh)')
        self.addMembertoCat('Building energy', 'Lights Misc. A (kWh)')
        self.addMembertoCat('Building energy', 'Equip Misc. H (kWh)')
        self.addMembertoCat('Building energy', 'Other Process (kWh)')
        self.addMembertoCat('Building energy', 'MEC elevators energy (kWh)')
        self.saveCatFile()


