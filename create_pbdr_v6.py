#!/usr/bin/python -tt

import html5lib
import sys
import os
import re
import time
import datetime

import numpy as np
import pandas as pd

def convert_date(x):
  if x is np.NaN:
    ret=x
  else:
    #cut the first 10 characters and format it as date
    ret=datetime.datetime.strptime(x[:10],'%m/%d/%Y').date()
  return ret
  #end 

def convert(filename):
  print('Converting -> ' + filename)
  #get the person number 
  file=open(filename)
  text=file.read()
  match=re.search(r'[\s]+([\d]+)</li>[\s]+</ul>',text)
  if match:
    pernum=match.group(1)
  else:
    pernum='per'
  file.close()  
  print('Person number-> ' + pernum)
  
  ctime=time.strftime("%Y%m%d_%H%M%S")
  xfname='pbdr_'+pernum+'_'+ctime+'.xlsx'
  
  data=pd.read_html(filename)
  writer=pd.ExcelWriter(xfname)
  workbook=writer.book
  num_format=workbook.add_format({'num_format': '#0'})

  #create a dictionary of dataframes 
  dfs={'per':data[2],
       'paf':data[10],
       'pps':data[12],
       'brn':data[14],
       'ptnl':data[26],
       'pil':data[28],
       'prtt':data[32],
       'prtt_rslt':data[34],
       'dpnt':data[38],
       'bnf':data[40],
       'actn':data[42],
       'pay':data[46]
  }

  sort_dict={ 'per':["Effective Start Date"],
       'paf':["Effective Start Date"],
       'pps':["Effective Start Date"],
       'brn':["Effective Start Date","Benefit Relation Name"],
       'ptnl':["Occurred"],
       'pil':["Occurred"],
       'prtt':["Enrollment Start"],
       'prtt_rslt':["PerInLer ID","Enrt Result ID"],
       'dpnt':["PerInLerID","EnrtResultID"],
       'bnf':["PerInLerId","Enrollment Result ID"],
       'actn':["PerInLerID","Enrollment Result ID"],
       'pay':["EnrollmentResultID","Effective Start Date"]
  }

  format_dict={ 'per': ["Effective Start Date","Effective End Date"]
               ,'paf':["Effective Start Date","Effective End Date"]
               ,'pps':["Effective Start Date"]
               ,'brn':["Effective Start Date","Effective End Date"]
               ,'ptnl':["Occurred","Detected"]
               ,'pil':["Occurred","Started"]
               ,'prtt':["Enrollment Start","Enrollment End","Original Start"]
               ,'prtt_rslt':["Rate Start","Rate End","LE Occured Date"]
               ,'dpnt':["Occured Date","Dpnt Cvg Start","Dpnt Cvg End"]
               ,'bnf':["Occurred Date","Dsgn Start","Dsgn End"]
               ,'actn':["Due Date"]
               ,'pay':["Effective Start Date","Effective End Date"]
  }

  for sheetname,df in dfs.items():
    print('Sheetname:',sheetname)
    #col_list=df.loc[0].tolist()
    #df.columns=col_list
    #df.drop(df.index[0],inplace=True)
    sort_key=sort_dict[sheetname]
    #change the data type of datetime
    try:
      for column_name in format_dict[sheetname]:
        df[column_name] = df[column_name].apply(lambda x: convert_date(x))
    except:
      print('Error formatting',sys.exc_info()[0])
    for column_name in df.columns.tolist():
      match=re.search(r'[\w]*id$',column_name.lower())
      if match:
        #the column name ends with id
        try:
          df[column_name]=df[column_name].fillna(0).astype('int64')
        except:
          print('Error formatting astype',sys.exc_info()[0])
    df.sort_values(by=sort_key,inplace=True)
    df.to_excel(writer,sheet_name=sheetname,index=False)
    worksheet=writer.sheets[sheetname]
    worksheet.freeze_panes(1,0)    
    for idx,col in enumerate(df):
      series=df[col]
      max_len=max((
        series.astype(str).map(len).max(),
        len(str(series.name))
        ))+1
      match = re.search(r'[\w]*id$',col.lower())
      try:
        if match:
        # the column is a number
          worksheet.set_column(idx,idx,max_len,num_format)
        else:
          worksheet.set_column(idx,idx,max_len)
      except:
        print('Error in setting excel number format',sys.exc_info()[0])
  #end for loop
  writer.save()    
  writer.close()

def main():
  #check for input file name
  if len(sys.argv) != 2:
    print('usage: ./create_pbdr.py filename')
    sys.exit(1)
  filename=sys.argv[1]
  dirname=os.getcwd()
  abspath = os.path.join(dirname,filename)
  if os.path.exists(abspath):
    convert(abspath)
  else:
    print('check the filename !!')


if __name__ == '__main__':
  main()
