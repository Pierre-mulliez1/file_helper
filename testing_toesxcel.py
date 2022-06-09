#!/usr/bin/env python
# coding: utf-8

# ## File helper
# ### Author: 
# Pierre Mulliez
# ### Creation date: 
# 26-10-2021
# ### Last modified: 
# 03-04-2022
# ### Description: 
# Helper function to convert csv to excel with the right encoding, Compare files for differences, Load manually ACCESS files.
# ### Contact: 
# Pierremulliez1@gmail.com

# In[5]:


#Small code chunk to format excel files 
import pandas as pd 
import openpyxl 
import os
from openpyxl import worksheet
import re

def to_excel(sep1= ',', 
             back = 'Y', 
             empty = 'Y', 
             apply_conditions_file = (),
             Date_start = '',
             Date_end = '',
             Opened = '',
             Purchase = '',
             Account = '',
            Duplicates = ''):
    #emptying the destination folder 
    get_ipython().system('del excels /q /s')
    
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
        df1 = pd.read_csv(DATA_PATH,sep = sep1) #, encoding='cp1252'
            
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
        with pd.ExcelWriter("excels/{}.xlsx".format(txt[0]),engine="openpyxl",
    if_sheet_exists="replace") as writer:
            df1.to_excel(writer, sheet_name="extraction_1",index = False)  
            if count in apply_conditions_file:
                print('Applying conditions on file {}'.format(count))
                d = {'Date_start': Date_start, 'Date_end': Date_end, 'Opened': Opened,
                     'Purchase': Purchase, 'Account': Account, 'Duplicates': Duplicates}
                df2 =  pd.DataFrame(data = d, index = [0])
                df2.to_excel(writer, sheet_name="Conditions",index = False)  
        
        get_ipython().system('echo %CD%')
        
        #Warning and escape 
        if count > 1:
            print('Converting multiple files do not input conditions')
        elif count > 8:
            print('Error, too many files to convert ')
            break 
        count += 1
    
    ##Create backup?##
    if (back == 'Y'):
        get_ipython().system('MOVE /Y data\\* backup')
        
    ##emptying the source folder
    if (empty == 'Y'):
        get_ipython().system('del data /q /s')


# In[10]:


to_excel(
         sep1 = ',', 
         apply_conditions_file = (1,20),
             Date_start = '',
             Date_end = '',
             Opened = 'YES',
             Purchase = 'Only role company identified and in the next 15 days of opening email for specific product only',
             Account = 'Communication auto LPS',
            Duplicates = 'Allowed for communication and contactability information, if an email has purchased different products or opened different link or the same link at different times'
            )


# In[ ]:





# In[3]:


#Small code chunk to compare files
import pandas as pd 
import openpyxl 
import os
from openpyxl import worksheet
import re
def compare_files(sep1= ',',back = 'Y',empty = 'Y'):
    try:
        files = os.listdir('data/')
        count = 1
        for fil in files:
            DIRECTORY_WHERE_THIS_FILE_IS = os.path.dirname(os.path.abspath(('data/'+fil)))
            DATA_PATH = os.path.join(DIRECTORY_WHERE_THIS_FILE_IS, fil)
        #read from source using delimiter and using the right encoding for spain 
            if count == 1:
                df1 = pd.read_csv(DATA_PATH,sep = sep1) #, encoding='cp1252'
                print('file 1 generated succesfully')
            elif count == 2:
                df2 = pd.read_csv(DATA_PATH,sep = sep1) #, encoding='cp1252'
                print('file 2 generated successfully')
            else:
                print('Made for 2 files only')
                break
            count += 1
    except:
        print('Delimiter likely not defined properly')
        
    ##Create backup?##
    if (back == 'Y'):
        get_ipython().system('MOVE /Y data\\* backup')
    
    ##emptying the source folder
    if (empty == 'Y'):
        get_ipython().system('del data /q /s')
        
    #bigger than prev file ?
    print('Compare file sizes:')
    if len(df2) >= len(df1):
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


# In[4]:


compare_files()


# In[5]:


#Small code chunk to get access files correctly 
import pandas as pd 
import os
import re

def access_f():
    #emptying the destination folder 
    get_ipython().system('del excels /q /s')
    
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

        df1 = pd.read_excel(DATA_PATH) #, encoding='cp1252'

        #generate correct output name
        txt0 = str(os.path.basename(DATA_PATH))
        txt = re.split('\.',txt0)
        accessttxt = 'ACCESS ' + txt[0]
        #omit row indexing
        df1.reset_index(drop=True, inplace=True)
        
        #find the user file 
        result = re.search('ACCESS COUNTRIES_EXPORTS_Users_.*',accessttxt)
        print(accessttxt)
        if result:
            #filter the right data
            df2 = df1[((df1['Sign-up origin + date'].str.contains('SAP')) 
                        & (df1['Fecha del último acceso del usuario'].str.contains('2')) |
                       ( df1['Sign-up origin + date'].str.contains('SAP',na=False) == False ) )]
            #Backup
            df3 = df1.to_excel("excels/{}.xlsx".format('BACK ' + accessttxt),index=False)
            df4 = df2.to_excel("excels/{}.xlsx".format(accessttxt),index=False)
        else:
            df2 = df1.to_excel("excels/{}.xlsx".format(accessttxt),index=False)
        get_ipython().system('echo %CD%')
        
        #Warning and escape 
        if count > 1:
            print('Converting multiple files ')
        elif count > 5:
            print('Error, too many files to convert ')
            break 

        count += 1
            
    #emptying source folder
    get_ipython().system('del data /q /s')


# In[6]:


access_f()


# In[ ]:




