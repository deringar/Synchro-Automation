# -*- coding: utf-8 -*-
"""
Created on Fri Sep 27 2024
Last modified on Thurs Oct 3 2024

@authors: philip.gotthelf, alex.dering - Colliers Engineering & Design
"""

from core import MainWindow, Base, read_input_file

# Step 2: V/C Ratio, LOS

if __name__ == '__main__':
    read_input_file("test-input.xlsx")
    # tfile = 'C:\\Users\\pgard\\Documents\\Synchro Automation\\synchronizer\\tests\\2020 EXISTING AM.txt'
    # import pprint
    # pprint.pprint(standardize(tfile))
    root = Base()
    # root.attributes('-topmost', True)
    root.resizable(True, True)
    # icon = tk.PhotoImage(file='Logo.png')
    # root.iconphoto(True, icon)
    app = MainWindow(root)
    root.main_win = app
    root.windows['main'] = app
    app.run()
