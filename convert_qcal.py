import os
import sys
import pandas as pd
import numpy as np

"""
Parse a .csv output file from qCal to a computer readable style in xls format.

Run as

    python convert_qcal.py qcal_file.csv

Output file is input_parsed.xls and saved in the same folder.


By: Emil Ljungberg, Lund, 2024    
"""

def str2arr(s):
    s = s.replace('Hyperfine, Inc.', 'Hyperfine Inc.')
    s = s.replace('\n', '')
    ss = s.split(',')
    ss = [x.replace('"',"") for x in ss]
    return ss

def parse_study_params(fname, protocol_number):
    with open(fname, 'r') as f:
        for l in f:
            s = str2arr(l)
            if s[0] == "Study Parameters":
                params = s[1:]
                l = f.readline()
                values = str2arr(l)[1:]
    
    D = {k:x for k,x in zip(params, values)}
    df = pd.DataFrame(D, index=[0])
    df['Protocol Number'] = protocol_number
    return df

def parse_ADC_voi(fname, protocol_number):
    with open(fname, 'r') as f:
        for l in f:
            s = str2arr(l)
            if s[0] == "ADC VOI Statistics":
                col_names = s[1:]

                # skip the units
                f.readline()
                
                # Read in all the vois
                VOIs = []
                for i in range(14):
                    VOIs.append(str2arr(f.readline())[1:])
 
    df = pd.DataFrame(columns=col_names)
    for i in range(len(VOIs)):
        df.loc[len(df)] = [np.nan if len(x)==0 else x for x in VOIs[i]]

    dtypes = [float]*len(col_names)
    dtypes[0] = int
    dtypes[1] = str
    dtypes[2] = int
    for i,c in enumerate(df.columns):
        print(i)
        df[c] = df[c].astype(dtypes[i])
    
    df['Protocol Number'] = protocol_number
    return df

def parse_T2_voi(fname, protocol_number):
    with open(fname, 'r') as f:
        for l in f:
            s = str2arr(l)
            if s[0] == "T2 Contrast VOI Statistics":
                col_names = s[1:]

                # skip the units
                f.readline()
                
                # Read in all the vois
                VOIs = []
                for i in range(28):
                    VOIs.append(str2arr(f.readline())[1:])
   
    df = pd.DataFrame(columns=col_names)
    for i in range(len(VOIs)):
        df.loc[len(df)] = [pd.NA if len(x)==0 else x for x in VOIs[i]]

    dtypes = [float]*len(col_names)
    dtypes[0] = int
    dtypes[1] = str
    dtypes[2] = int
    for i,c in enumerate(df.columns):
        df[c] = df[c].astype(dtypes[i])
    
    df['Protocol Number'] = protocol_number
    return df

def parse_temp(fname, protocol_number):
    finished = False
    with open(fname, 'r') as f:
        while not finished:
            l = f.readline()
            s = str2arr(l)
            if s[0] == "Additional Calculations":
                l2 = f.readline().split('","')
                temps = []
                temp_labels = []
                for i in range(len(s)):
                    if 'Temperature' in s[i]:
                        temps.append(float(l2[i]))
                        temp_labels.append(s[i])
                finished = True
    
    df = pd.DataFrame(columns=['Protocol Number'] + temp_labels)
    df.loc[len(df)] = [int(protocol_number)] + temps
    df['Protocol Number'] = df['Protocol Number'].astype(int)
    return df


def convert_to_xls(fname, out_fname, protocol_number):
    info = parse_study_params(fname, protocol_number)
    print(info)
    ADC = parse_ADC_voi(fname, protocol_number)
    ADC['Study Date'] = info['Study Date'][0]
    print('Parsed ADC')
    
    temps = parse_temp(fname, protocol_number)
    temps['Study Date'] = info['Study Date'][0]
    print('parsed temps')

    t2voi = parse_T2_voi(fname, protocol_number)
    t2voi['Study Date'] = info['Study Date'][0]
    print('Parsed t2vi')
    writer = pd.ExcelWriter(out_fname)

    ADC.to_excel(writer, sheet_name='ADC', index=False)
    info.to_excel(writer, sheet_name='info', index=False)
    temps.to_excel(writer, sheet_name='temperature', index=False)
    t2voi.to_excel(writer, sheet_name='T2w', index=False)

    # Save the Excel file
    writer.close()


def main():
    fname = sys.argv[1]
    out_fname = sys.argv[2]
    protocol_number = int(sys.argv[3])
    convert_to_xls(fname, out_fname, protocol_number)
    
if __name__ == '__main__':
    main()