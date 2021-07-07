from tkinter import filedialog
from tkinter import *
from datetime import *
import pandas as pd
import numpy as np

def bom_check(files):
    boms = []  
    for bb in files:
        boms.append(bb.get()) #files is dic...
    df = pd.read_excel(boms[0], sheet_name = 'BOMSheet1', usecols = 'C, H, J')
    df = df[5:]
    new_df = pd.DataFrame(df)
    new_df = new_df.rename(columns = {'Unnamed: 2': 'Number', 'Unnamed: 7': 'Quantity', 'Unnamed: 9': 'Reference Designators'}, inplace=False)
    new_df = new_df.fillna(1) #fill NA/NaN value to 1
    check_list = []
    for index, row in new_df.iterrows():
        pn = row['Number']
        pcs = row['Quantity']
        reference = row['Reference Designators'] 
        pcs = float(pcs)
        ref = str(reference)
        ref = ref.split(',')
        pcs_ref = len(ref)
        if pcs_ref == pcs:
            continue
        else:
            c = [pn, pcs, pcs_ref]
            check_list.append(c)
            continue
    pcs_ck = len(check_list)
    ck = pd.DataFrame(check_list, columns = ['Part Number', 'Quantity', 'Number of Reference Designators'])
    if pcs_ck == 0:
        print('This BOM is perfect!')
    else:
        print("Those part_number's quantity is wrong as below")
        print('----------------------------------------------')
        print(ck)

def openfile(ent):
    file_in = filedialog.askopenfilename(filetypes = (("Excel File","*.xlsx .xls"),("all files","*.*")))
    ent.insert(0, file_in) 

def makeform(root, feilds):
    rows = []
    entries = [];
    for ff in feilds:
        row = Frame(root)
        lab = Label(row, width=12, text=ff)
        ent = Entry(row, width=80)
        row.pack(side=TOP, fill=X)
        lab.pack(side=LEFT)     
        ent.pack(side=LEFT, fill=X)
        entries.append(ent)
        rows.append(row)        
    Button(rows[0], text='SEL', command= lambda:openfile(entries[0])).pack(side=RIGHT)
    return entries

def main():
    bom_ck = Tk()
    bom_ck.title('BOM check')
    BOM_IN = ['BOM_IN']
    BOM = makeform(bom_ck, BOM_IN)
    BOM_CHECK = Button(bom_ck, text = 'BOM check', fg='#000079', bg='#66B3FF', font=('Arial', 12), 
                      command= lambda:bom_check(BOM)).pack(side = BOTTOM)
    bom_ck.mainloop()

if __name__ == "__main__":
    main();