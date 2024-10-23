from enum import Enum

class Cell:
    def __init__(self, id, position_x, position_y, type, sheetname):
        self.id = id
        self.position_x = position_x
        self.position_y = position_y
        self.type = type
        self.sheetname = sheetname
