# -*- coding: utf-8 -*-
"""
Last modified on Fri Sep 27 2024

@authors: philip.gotthelf, alex.dering
"""

from core import MainWindow, Base

# Step 2: V/C Ratio, LOS

if __name__ == '__main__':
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
