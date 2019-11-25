# This module contains Category class defining categories used in qtreewidget to categorise attributes of dataframe
class Category:
    def __init__(self, name, mutable = True):       # constructor
        self.name = name            # category name
        self.members = list()       # list of its attribute members
        self.mutable = mutable

    def setName(self, name):        # setter function for its name
        self.name = name

    def addMembers(self,member):    # member appending function
        self.members.append(member)