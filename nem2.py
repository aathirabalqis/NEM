import os
import math
import shutil
import numpy as np
import pandas as pd 
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.scrolledtext import ScrolledText
from openpyxl import load_workbook

root = Tk()

root.title('NEM')
root.geometry('500x390') 

def nemfile():

    global nempath, newnem, folder

    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    Label(root, text = filenames[0]).place(x = 175, y  = 22) 
    nempath = filenames[0] 
    
    folder = '\\'.join(nempath.split('/')[:-1])
    title = nempath.split('/')[-1].split('.')[0]
    
    # in case XLSX
    newnem = folder + '\\' + title + '.xlsx'
    print(nempath)
    print(newnem)
    
def atbfile():
    
    global atbdf, atbpath, folder, out, month

    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    Label(root, text = filenames[0]).place(x = 175, y  = 62) 
    atbpath = filenames[0] 
    
    month = atbpath.split('/')[-1].split('.')[0].split(' ')[-1]
    out = folder + '\\Reading NEM ' + month + '.xlsx'
    
    # read atb file
    atbdf = pd.read_excel(atbpath)
    atbdf = atbdf.rename(columns = lambda x: x.strip())    
    atbdf['Autobill'] = atbdf.loc[:,'Sec.Obj.Ky']
    print(atbdf['Autobill'])
    atbdf = atbdf.rename(columns = {'Sec.Obj.Ky' : 'Device No.'})
    atbdf = atbdf[['Device No.','Autobill']]
    print(atbdf)

def devstat():

    global dsdf, out, all_ids, ind, flag
    
    b["state"] = DISABLED
    b1["state"] = DISABLED
    
    ind = 0
    new_ids = []
    all_ids = []

    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    Label(root, text = filenames[0]).place(x = 175, y  = 102) 
    dspath = filenames[0]  

    dsdf = pd.read_excel(dspath)
    dsdf = dsdf[['ID','Device Status']]
    dsdf = dsdf.rename(columns = {'ID' : 'Device No.'})
    
    df = pd.read_excel(out)
    df = pd.merge(left = df, right = dsdf, how = 'left')
    print(df['Reading Status'])
    
    df.loc[(df['Reading Status'] != 'VAL') & (df['Device Status'] != 'Commissioned'), 'Reading Status'] = 'Meter not Connected to Network'
    df.loc[(df['Reading Status'] != 'VAL') & (df['Device Status'] == 'Commissioned'), 'Reading Status'] = 'Meter not Reporting'
    df = df.drop(columns = ['Device Status'])
    
    udf = df.loc[(df['Reading Status'] == 'Meter not Reporting'), 'Device No.']

    # output ID to get val or not
    new_ids = list(filter(None, udf))
    act_total.config(text = 'Total IDs = ' + str(len(new_ids)))
    all_ids.append(new_ids)
    
    if len(new_ids) > 1000:
        flag = 'SQL'
        divide(new_ids, 1000)
    
    textbox.delete('1.0', END)
    textbox.insert(END,'\'' + '\',\''.join(all_ids[ind]) + '\'')
    
    page.config(text = '1 of ' + str(len(all_ids)))
    total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind]))) 
    outp.config(text = 'Find UDC ID')
    
    os.remove(out)
    df.to_excel(out, index = False)
    
def getrep():

    global atbdf, nempath, all_ids, ind, flag, out, newnem
    
    b["state"] = DISABLED
    b1["state"] = DISABLED
    
    ind = 0
    new_ids = []
    all_ids = []

    # copy nem file - for .XLSX
    shutil.copy(nempath, 'temp.xlsx')
   
    # read nem file
    writer = pd.ExcelWriter('temp.xlsx',  mode = 'a', engine = 'openpyxl')
    nemdf = pd.read_excel('temp.xlsx') 
    
    nemdf = nemdf.loc[(nemdf['Portion'] == 'NORMAL31') | (nemdf['Portion'].str.startswith('SPOT'))]
    print(nemdf)

    # merge and write 
    if 'Autobill' not in nemdf.columns:
        print('already')
        os.remove(nempath)
        nemdf = pd.merge(left = nemdf, right = atbdf, how = 'left')
        nemdf.loc[(pd.isnull(nemdf['Autobill'])), 'Autobill'] = '#N/A'
        nemdf.to_excel(newnem, index = False)
    
    writer.close()
    os.remove('temp.xlsx')
    print('done part 1')
    
    # create new file for read
    if not os.path.exists(out):
        nadf = nemdf.loc[(nemdf['Autobill'] == '#N/A')|(pd.isnull(nemdf['Autobill']))]
        nadf = nadf[['State','Station','Station Description','Installation','Contract Acc.','Customer Name','Device No.','Portion']]
        nadf.to_excel(out, index = False)
    
    else:
        nadf = pd.read_excel(out)
    
    print(nadf)
    
    # output ID to get val or not
    new_ids = list(filter(None, nadf['Device No.']))
    act_total.config(text = 'Total IDs = ' + str(len(new_ids)))
    all_ids.append(new_ids)
    
    if len(new_ids) > 1000:
        flag = 'SQL'
        divide(new_ids, 1000)
    
    textbox.delete('1.0', END)
    textbox.insert(END,'\'' + '\',\''.join(all_ids[ind]) + '\'')
    
    page.config(text = '1 of ' + str(len(all_ids)))
    total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))
    outp.config(text = 'Find Readings')
    
def getnr():
    print('nr')
    
    global out, all_ids
    
    b["state"] = DISABLED
    b1["state"] = DISABLED
    
    inpu = textvar.get()
    textvar.set('')
    row = inpu.split('\n')
    
    big = []
    sets = []
    all_ids = []
    
    df = pd.read_excel(out)
    
    if inpu:
        for i in row:
            temp = i.split('\t') 
            temp[2] = int(float(temp[2]))
            big.append(temp)
        
        dfexp = pd.DataFrame(big, columns = ['METER_ID','MEAS_TYPE_ID', 'READ_VALUE', 'READ_TIME', 'VAL_STATUS','LAST_UPD_TIME'])
        datee = dfexp.loc[1,'READ_TIME'].split(' ')[0]
        dfexp = dfexp[['METER_ID','MEAS_TYPE_ID', 'READ_VALUE', 'VAL_STATUS']]
        
        for i in range(len(dfexp)):
            if dfexp['MEAS_TYPE_ID'][i] not in sets:
                sets.append(dfexp['MEAS_TYPE_ID'][i])
                
                dfs = dfexp.loc[(dfexp['MEAS_TYPE_ID'] == dfexp['MEAS_TYPE_ID'][i]), ('METER_ID','READ_VALUE')]
                dfs = dfs.rename(columns = {'METER_ID' : 'Device No.', 'READ_VALUE':str(dfexp['MEAS_TYPE_ID'][i])}) #,'VAL_STATUS':'Reading Status'})
                # print(dfs)
                df = pd.merge(left = df, right = dfs, how = 'left', on = 'Device No.')
        print(df)
        if not 'kWh Received (Register 01)' in df.columns:
            print('pls')
            df['MR Date'] = datee
            df['MR Date'] = pd.to_datetime(df['MR Date']).dt.date
            df = df[['State','Station','Station Description','Installation','Contract Acc.','Customer Name','Device No.','Portion','MR Date','106','110','113','105']]
            df.loc[(pd.isnull(df['106']) == False), 'Reading Status'] = dfexp['VAL_STATUS']
            df.loc[(pd.isnull(df['106'])), ['106','110','113','105','Reading Status']] = '#N/A'
            
            df = df.rename(columns = {'106' : 'kWh Received (Register 01)', '110':'kW Demand (Register 09)','113':'kVARh Received (Register 11)','105':'kWh Delivered (Register 51)'})

        else:
            print('yes')
            df.loc[(pd.isnull(df['106'])), ['106','110','113','105','Reading Status']] = '#N/A'
            df.loc[(pd.isnull(df['MR Date'])), 'MR Date'] = datee
            df.loc[(pd.isnull(df['kWh Received (Register 01)'])), 'kWh Received (Register 01)'] = df['106']
            df.loc[(pd.isnull(df['kW Demand (Register 09)'])), 'kW Demand (Register 09)'] = df['110']
            df.loc[(pd.isnull(df['kVARh Received (Register 11)'])), 'kVARh Received (Register 11)'] = df['113']
            df.loc[(pd.isnull(df['kWh Delivered (Register 51)'])), 'kWh Delivered (Register 51)'] = df['105']
            
            
            df.loc[(df['kWh Received (Register 01)'] != '#N/A'), 'Reading Status'] = 'VAL'
            df = df.drop(columns = ['106','110','113','105'])

        os.remove(out)
        df.to_excel(out, index = False)  
        
        print('yay') 

        # output non-val ID
        nv = df.loc[(df['Reading Status'] != 'VAL'), 'Device No.']
        
        new_ids = list(filter(None, nv))
        act_total.config(text = 'Total IDs = ' + str(len(new_ids)))
        stri = str(','.join(new_ids))
        
        textbox.delete('1.0', END)
        textbox.insert(END, stri)
        total.config(text = '')
        outp.config(text = 'Find Device Status')
        
    else: # just to get NR meters in SQL
        nv = df.loc[(df['Reading Status'] == 'Meter not Reporting'), 'Device No.']
        new_ids = []
        
        new_ids = list(filter(None, nv))
        act_total.config(text = 'Total IDs = ' + str(len(new_ids)))
        all_ids.append(new_ids)
    
        if len(new_ids) > 1000:
            flag = 'SQL'
            divide(new_ids, 1000)
        
        textbox.delete('1.0', END)
        textbox.insert(END,'\'' + '\',\''.join(all_ids[ind]) + '\'')
        
        page.config(text = '1 of ' + str(len(all_ids)))
        total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))
        print('o')

def getudc():
    print('udc')
    
    global all_ids, ind, flag
    
    b["state"] = DISABLED
    b1["state"] = DISABLED
    
    ind = 0
    new_ids = []
    all_ids = []
    data = []
    
    inpu = textvar.get()
    textvar.set('')
    row = inpu.split('\n')
    
    for i in row:
        temp = i.split('\t')[1]
        print(temp)
        data.append(temp)
        
    df = pd.DataFrame(data, columns = ['UDC ID'])
    print(df)
    
    # output UDC IDs
    new_ids = list(filter(None, df['UDC ID']))
    act_total.config(text = 'Total IDs = ' + str(len(new_ids)))
    all_ids.append(new_ids)
    
    if len(new_ids) > 12:
        flag = 'UDC'
        divide(new_ids, 12)
    
    textbox.delete('1.0', END)
    textbox.insert(END, str(','.join(all_ids[ind])))
    
    page.config(text = '1 of ' + str(len(all_ids)))
    total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))
    
def split():

    global out, folder, month
    
    df = pd.read_excel(out)
    df['NewState'] = df['State']
    df.loc[(df['State'] == 'PAH') | (df['State'] == 'TRE') | (df['State'] == 'KEL'), 'NewState'] = 'EAST'
    
    states = [] 
    states = df['NewState'].drop_duplicates().tolist()
    
    print(states)
    
    for i in states:
        tempdf = df.loc[(df['NewState'] == i)]
        
        tempdf = tempdf.drop(columns = ['NewState'])
        
        outt = folder + '//Reading NEM ' + str(i) + ' ' + month + '.xlsx'
        tempdf.to_excel(outt, index = False)




def test():
    inpu = textvar.get()
    textvar.set('')
    row = inpu.split('\n')
    
    if len(new_ids) > 12:
        flag = 'UDC'
        divide(new_ids, 12)
    
    textbox.delete('1.0', END)
    textbox.insert(END, str(','.join(all_ids[ind])))
    
    total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))
    
def divide(array,lim):

    global all_ids, ind

    num = math.ceil(len(array)/lim)
    all_ids = []
    for i in range(num):
        temp = array[lim*i:lim*(i+1)]
        print(len(temp))
        all_ids.append(temp)
    
    b["state"] = NORMAL   
    return all_ids
    
def nextt(): 

    global all_ids, ind, flag

    b["state"] = DISABLED
    b1["state"] = NORMAL
    ind += 1
    
    textbox.delete('1.0', END)
    if flag == 'UDC':
        textbox.insert(END, str(','.join(all_ids[ind])))
    
    else:
        textbox.insert(END,'\'' + '\',\''.join(all_ids[ind]) + '\'') 
    
    if ind < len(all_ids)-1:
        b["state"] = NORMAL
        
    total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))
    page.config(text = str(ind+1) + ' of ' + str(len(all_ids)))
        
def back():

    global all_ids, ind, flag
    
    b["state"] = NORMAL
    b1["state"] = DISABLED
    ind -= 1
    
    textbox.delete('1.0', END)
    if flag == 'UDC':
        textbox.insert(END, str(','.join(all_ids[ind])))   
    else:
        textbox.insert(END,'\'' + '\',\''.join(all_ids[ind]) + '\'')
    
    if ind != 0:
       b1["state"] = NORMAL 
       
    total.config(text = 'Displayed IDs = ' + str(len(all_ids[ind])))
    page.config(text = str(ind+1) + ' of ' + str(len(all_ids)))

def exitt(event):
    root.quit()

Button(root, text = 'Total NEM', command = nemfile).place(x = 53, y = 20) 
Button(root, text = 'Autobill File', command = atbfile).place(x = 45, y = 60)
Button(root, text = 'Device Status', command = devstat).place(x = 40, y = 100)

Label(root, text = 'Path:').place(x = 140, y  = 22) 
Label(root, text = 'Path:').place(x = 140, y  = 62) 
Label(root, text = 'Path:').place(x = 140, y  = 102) 

textvar = StringVar()
Entry(root, textvariable = textvar, width = 26, font = ('calibre',10,'normal')).place(x = 40, y = 150, height = 20)

Button(root, text = 'NEM Rep.', command = getrep).place(x = 235, y = 147)
Button(root, text = 'NR Meters', command = getnr).place(x = 299, y = 147)
Button(root, text = 'UDC ID', command = getudc).place(x = 366, y = 147)
Button(root, text = 'Split', command = split).place(x = 416, y = 147)

textbox = ScrolledText(root, height = 8, width = 51)
textbox.place(x = 40, y = 210)

Label(root, text = 'SAB').place(x = 470, y = 370) 

b = Button(root, text = 'Next', command = nextt)
b.place(x = 77, y = 345)
b1 = Button(root, text = 'Back', command = back)
b1.place(x = 40, y = 345)

b["state"] = DISABLED
b1["state"] = DISABLED

total = Label(root, text = '')
total.place(x = 350, y = 347)
act_total = Label(root, text = '')
act_total.place(x = 115, y = 347)
page = Label(root, text = '')
page.place(x = 410, y = 187)
outp = Label(root, text = 'Output')
outp.place(x = 40, y = 187)

root.resizable(True, False) 
root.bind('<Escape>', exitt)
root.mainloop()