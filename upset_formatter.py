# -*- coding: utf-8 -*-
"""
Created on Thu Feb  2 15:23:41 2023

@author: KA864PA
"""

import pandas as pd
import os
import openpyxl
import json
datapath = os.getcwd()+"\\data\\"


def load_excel(file_name, dp_other = None, verbose = False, skip_rows = 0):
    """
    For all valid and visible sheets, get dict of sheetname :[columns, data]
    """
    f_dir = datapath+file_name if dp_other is None else file_name
    #get visible and valid sheets
    visible_sheets = [ws.title for ws in openpyxl.load_workbook(f_dir,read_only=True).worksheets 
        if ws.sheet_state != 'hidden' ]
    #print(visible_sheets)
    
    #get cols and rows
    df_sheet_map = {i:[v.columns.values.tolist(),v]\
        for i,v in pd.read_excel(f_dir, sheet_name=None,skiprows=skip_rows).items()\
        if i in visible_sheets}
    # print(df_sheet_map['Sheet2'])
    return df_sheet_map
    pass
    

def get_unique_vals(df,col):
    return list(set(df[col].values.tolist()))
    #taking dataframe, get unique values as list.
    pass

def concat_counts(_dict, delimiter= ";"):
    return delimiter.join([delimiter.join(list(map(lambda x: str(x), v.values()))) for k,v in _dict.items()])
    pass
def build_csv(main_field, subfield_dict, upset_dict):
    header_row = main_field+ ";"+";".join([";".join(list(map(lambda x:"{} ({})".format(x,k), v))) for k,v in subfield_dict.items()])
    print(header_row)
    # print(upset_dict)
    body = "\n".join([k+";"+concat_counts(v) for k,v in upset_dict.items()])
    print(body)
    # return header_row +"\n"+body
    with open(datapath+"upset_graph.csv", "w") as text_file:
       text_file.write(header_row +"\n"+body)
    pass

def build_json(subfield_len):
    json_dict_template = {
    	"file": "https://raw.githubusercontent.com/subledgerdev/set_graph/main/data/upset_graph.csv",
    	"name": "Subledger ",
    	"header": 0,
    	"separator": ";",
    	"skip": 0,
    	"meta": [
    		{ "type": "id", "index": 0, "name": "Name" }
    	],
    	"sets": [
    		{ "format": "binary", "start": 1, "end": subfield_len }
    	]
    }
    with open(datapath+"upset_graph.json", "w") as outfile:
        json.dump(json_dict_template, outfile)
    pass

def format_ledger(main_field,pivot_fields, main_sheet = 'Sheet2',_file = 'upset.xlsx', verbose = False):
    #get dicts
    main_col_dict = {}
    #get uniques in col i.e., left side is acctg basis
    
    
    sheet_dict = load_excel(_file)
    if main_sheet not in sheet_dict.keys():
        raise Exception("sheet not found: {}".format(main_sheet))
        return
    
    ledger_cols, ledger_data = sheet_dict[main_sheet]
    
    err_cols = " ".join([i for i in pivot_fields + [main_field] if i not in ledger_cols])
    if err_cols:
        raise Exception("desired columns not extant in spreadsheet: {}".format(err_cols))
    
    #get empty upset graph matrix
    subfields = {k:get_unique_vals(ledger_data,k) for k in pivot_fields} #lazyyyyy code repetition
    # print(sum([len(i) for i in subfields.values()]))
    
    upset_matrix = {i:
        {k:{subfield:0 for subfield in get_unique_vals(ledger_data,k)} for k in pivot_fields}
                     for i in get_unique_vals(ledger_data,main_field)}
    for k,v in upset_matrix.items():
        for kk,vv in v.items():
            for kkk, vvv in vv.items():
                query = len(ledger_data[(ledger_data[main_field]==k) &  (ledger_data[kk]==kkk)])
                upset_matrix[k][kk][kkk] =  query# len(ledger_data[(ledger_data[main_field]==k) &  (ledger_data[kk]==kkk)])
                if verbose:
                    print("{} | {} | {}".format(k,kk,query))
                    
    build_csv(main_field,subfields,upset_matrix)
    build_json(sum([len(i) for i in subfields.values()]))
        # print("=======================================")
    # print(main_col_dict)
    
    
    # df2 = len(df[(df["Courses"]=="Pandas") & 
    #          (df["Fee"]==35000) & 
    #          (df["Duration"]>= "35days")])
                     # }
    # asdf = set(ledger_data[main_field].values.tolist())
    # print(main_col_dict)
    # 
        
# {'ledger_accounting_basis': {'GAAP Adjustment': 0, 'Common': 0, 'Nominal Event': 0}, 
#  'ledger_type': {'Money': 0, 'Position': 0, 'Adjustment': 0}
    
    pass
    
if __name__ == '__main__':
    # wing_wong_ray_wong = {'ledger_accounting_basis': {'GAAP Adjustment': 0, 'Common': 1, 'Nominal Event': 0}, 'ledger_type': {'Adjustment': 0, 'Money': 1, 'Position': 0}}
    # print(concat_counts(wing_wong_ray_wong))
    format_ledger('ledger_name',['ledger_accounting_basis','ledger_type'], verbose = False)
    
    pass