# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import xlwings as xw
import pandas as pd

class Manipulator:
    def __init__(self,book):
        self.wb = xw.Book(book)
        self.sht = self.wb.sheets("Risk Profiles")

    def test(self):
        data = self.wb.selection
        print (data.value)


def main():
    # wb = xw.Book("book1.xlsx")

    # wb = xw.Book()

    # sht = wb.sheets['Sheet1']
    # sht.range('A1').value = "Hello"
    
    # mp = Manipulator("BTH.xlsx")
    # mp.test()
    
    df = pd.read_excel('BTH.xlsx',"Risk Profiles")
    print(df)
    
    
if __name__ == "__main__":
    main()