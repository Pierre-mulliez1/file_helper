#Author: Pierre Mulliez
#Creation date: 26-10-2021

  #Small code chunk to format excel files 
import pandas as pd 
import openpyxl 
import os
from openpyxl import worksheet
import re

def to_excel(sep1= ',', back = 'Y', empty = 'Y'):
    #get the name of the file in folder 
    files = os.listdir('data/')
    fil = ''
    count = 1
    #Convert multiple files 
    for f in files:
        fil = f
        DIRECTORY_WHERE_THIS_FILE_IS = os.path.dirname(os.path.abspath(('data/'+fil)))
        DATA_PATH = os.path.join(DIRECTORY_WHERE_THIS_FILE_IS, fil)
        #read from source using delimiter and using the right encoding for spain 
        df1 = pd.read_csv(DATA_PATH,sep = sep1, encoding='cp1252')

        if len(df1.columns) < 2:
            print('WARNING, only one collumn found')
        elif len(df1.columns) < 3:
            print('WARNING, only two collumn found')

        ##set password ?##
        #df1 = df1.worksheets[0]
        #df1.protection.set_password('test')

        #generate correct output name
        txt0 = str(os.path.basename(DATA_PATH))
        txt = re.split('\.',txt0)
        #omit row indexing
        df1.reset_index(drop=True, inplace=True)
        df2 = df1.to_excel("excels/{}.xlsx".format(txt[0]),index=False)
       
        #Warning and escape 
        if count > 1:
            print('Converting multiple files ')
        elif count > 3:
            print('Error, too many files to convert ')
            break 
        count += 1

    
def compare_files(file1,file2,sep1= ','):
    DIRECTORY_WHERE_THIS_FILE_IS = os.path.dirname(os.path.abspath(('data/'+file1)))
    DATA_PATH = os.path.join(DIRECTORY_WHERE_THIS_FILE_IS, file1)
    DATA_PATH2 = os.path.join(DIRECTORY_WHERE_THIS_FILE_IS, file2)
    try:
        df1 = pd.read_excel(DATA_PATH)
        df2 = pd.read_excel(DATA_PATH2)
    except:
        print('Delimiter likely not defined properly')
        
    #bigger than prev file ?
    print('Compare file sizes:')
    if len(file2) >= len(file1):
        print('file2 bigger by {}'.format(len(df2) - len(df1)))
    else:
        print('file1 bigger by {}'.format(len(df1) - len(df2)))
    print('')
    
    #are the collumns name equal ?
    print('Collumn header equal ?')
    print(df1.columns == df2.columns)
    print('')
    
    #null proportion - first col
    print('Checking null values')
    print('Count of nulls - file1 col1 {}'.format(df1.iloc[:,0].isnull().sum()))
    print('Count of nulls - file2 col1 {}'.format(df2.iloc[:,0].isnull().sum()))

