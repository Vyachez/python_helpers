'''
Excel formatting module v.1

October 2019
V.Nesterov

'''
# importing modules
import numpy as np
import pandas as pd

import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting import Rule
from openpyxl.styles import Font, PatternFill, Border, NamedStyle, Side, Alignment
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
import string

def decor(path, file_name, tab_name, dataframe=[],
         row_h=None, col_ws=[20],
         header_h=30, conditional_coloring=None, exact_condition=True,
         header_fill_color="DDDDDD", header_text_color='000000', body_color="FFFFFF",
         left_col_color="FFFFFF",
         footnote=None):
    '''
        Function to format Excel spreadsheet. 
        .csv .xls or .xlsx format supported
        Single tab supported
        Takes:
            - path - full path to working directory (followed by '/');
            - file_name - target file name (including extension)
                or desired file name if dataframe passed;
            - tab_name - desired tab name;
            - dataframe - pandas dataframe to work with (if not with file) - optional;
            - row_h - row height, default=None - optional;
            - col_ws - columns width as a list of integers for each column in sequence,
                default=[20] - optional;
            - header_h - top row height, default=30 - optional;
            - conditional_coloring - dictionary where key is incell string to find and format and 
                value is color in HEX format (as string), default=None - optional;
            - condition - boolean to exactly match keys in conditional formatting dictionary or contain the key in cell
            - header_fill_color - color of top row in HEX (string), default="DDDDDD" (grey) - optional;
            - header_text_color - color of text in top row in HEX (string), default="000000" (black) - optional;
            - body_color - color of background in HEX (string), default="FFFFFF" (white) - optional;
            - left_col_color - color of left column in HEX (string), default="FFFFFF" (white) - optional;
            - footnote - footnote at the end of the document, default=None - optional.
        Returns:
            None. Saves spreadhseet in same directory
    '''
    if path[-1] != "/":
        path = path+"/"
    
    # if converting file itself
    if len(dataframe) == 0:
        # opening file
        if ".xls" in file_name or ".xlsx" in file_name:
            data = pd.read_excel(path+file_name, header=0, dtype=object)
        elif ".csv" in file_name:
            data = pd.read_csv(path+file_name, dtype=object)
        else:
            print("Error: Wrong data format. Please supply only .csv .xls or .xlsx format.")
            return None
    else:
        # reading just dataframe
        data = dataframe
        
    # dimensions
    rows_no, cols_no = data.shape
    rows_no+=1
    
    # #### Initializing openpyxl and setting up data
    # creating openpyxl object to read data in excel
    wb = Workbook()
    
    # defining tab
    main_tab = wb.active
    main_tab.title = tab_name
    
    # filling with data
    for r in dataframe_to_rows(data, index=False, header=True):
        main_tab.append(r)  
    
    # #### Formatting
    print("Formatting \n")
    # ##### defining formatting styles
    # body style 
    def add_body_style(wb):
        name = 'body'
        st = NamedStyle(name=name)
        st.font = Font(name='Calibri', bold=False, size=11)
        bd = Side(style='thin', color="000000")
        st.border = Border(left=bd, top=bd, right=bd, bottom=bd)
        st.alignment=Alignment(horizontal='left',
                            vertical='center',
                            text_rotation=0,
                            wrap_text=True,
                            shrink_to_fit=False,
                            indent=0)
        st.fill = PatternFill(start_color=body_color,
                           end_color=body_color,
                           fill_type='solid')
        wb.add_named_style(st)
        return name
    
    # header style
    def add_head_style(wb):
        name = 'headstyle'
        st = NamedStyle(name=name)
        st.font = Font(name='Calibri', bold=True, color=header_text_color, size=10)
        bd = Side(style='thin', color="000000")
        st.border = Border(left=bd, top=bd, right=bd, bottom=bd)
        st.alignment=Alignment(horizontal='center',
                            vertical='center',
                            text_rotation=0,
                            wrap_text=True,
                            shrink_to_fit=False,
                            indent=0)
        st.fill = PatternFill(start_color=header_fill_color,
                           end_color=header_fill_color,
                           fill_type='solid')
        wb.add_named_style(st)
        return name
       
    # left column style
    def add_leftcol_style(wb):
        name = 'indexer'
        st = NamedStyle(name=name)
        st.font = Font(name='Calibri', bold=False, size=10)
        bd = Side(style='thin', color="000000")
        st.border = Border(left=bd, top=bd, right=bd, bottom=bd)
        st.alignment=Alignment(horizontal='left',
                            vertical='center',
                            text_rotation=0,
                            wrap_text=True,
                            shrink_to_fit=False,
                            indent=0)
        st.fill = PatternFill(start_color=left_col_color,
                           end_color=left_col_color,
                           fill_type='solid')
        wb.add_named_style(st)
        return name
    
    # default columns dimensions             
    if cols_no != len(col_ws):
        print("Setting up default column width...\n"+\
              "If you want to specify width for each column, please \n"+\
              "specify it in col_ws list in order as columns follow.")
        col_ws = [20]*cols_no
    
    # ##### applying styles 
    # appending styles to workbook
    letters = list(string.ascii_uppercase)[:cols_no]
    rows = rows_no + 1
    
    # left column style
    index_style = add_leftcol_style(wb)
    for rw in range(2, rows):
        main_tab['A'+str(rw)].style = index_style
        
    # header style
    head_style = add_head_style(wb)
    for l in letters:
        main_tab[l+"1"].style = head_style
    
    # body style
    body_style = add_body_style(wb)
    for rw in range(2, rows):
        for l in letters[1:]:
            main_tab[l+str(rw)].style = body_style
                
    # ##### rows and columns dimensions
    # applying rows dimensions
    # header height
    main_tab.row_dimensions[1].height = header_h
    
    # regular row height
    if row_h != None:
        for dim in range(2, rows):
            main_tab.row_dimensions[dim].height = row_h
            
    # applying columns dimensions
    # iterating through number of columns
    for dim, w in zip(range(cols_no), col_ws):
        main_tab.column_dimensions[letters[dim]].width = w
    
    # ##### conditional formatting
    # adds conditional format to selected range - exact match
    def add_cond_text_format_exact(ws, text, color, start, end):
        '''
        Takes:
        - ws - worksheet object
        - text - as string
        - color - hex color
        - start cell+col string
        - end cell+col string
        '''
        fill = PatternFill(bgColor=color)
        dxf = DifferentialStyle(fill=fill)
        rule = Rule(type="cellIs", operator="equal", dxf=dxf)
        rule.formula = ['"{}"'.format(text)]
        ws.conditional_formatting.add(start+":"+end, rule)
    
    # adds conditional format to selected range - contains text
    def add_cond_text_format_contains(ws, text, color, start, end):
        '''
        Takes:
        - ws - worksheet object
        - text - as string
        - color - hex color
        - start cell+col string
        - end cell+col string
        '''
        print("using non-exact cond formatting")
        fill = PatternFill(bgColor=color)
        dxf = DifferentialStyle(fill=fill)
        rule = Rule(type="containsText", operator="containsText", text=text, dxf=dxf)
        rule.formula = ['NOT(ISERROR(SEARCH("{}",A2)))'.format(text)]
        ws.conditional_formatting.add(start+":"+end, rule)
    
    # inserting conditional formatting formula for ratings
    if conditional_coloring != None:
        if exact_condition:
            for val in conditional_coloring:
                add_cond_text_format_exact(main_tab, val, conditional_coloring[val], 'A2', letters[-1]+str(rows_no))
        else:
            for val in conditional_coloring:
                add_cond_text_format_contains(main_tab, val, conditional_coloring[val], 'A2', letters[-1]+str(rows_no))
    
    # ##### making filters
    print("Finalizing \n")
    # filtering
    main_tab.auto_filter.ref = "A1:"+letters[-1]+"1"
    
    # ##### other useful stuff
    # hiding gridlines
    main_tab.sheet_view.showGridLines = False
    
    # putting reference at the end of document
    if footnote != None:
        main_tab['A'+str(rows_no+2)].value = '* '+footnote
        
    # Saving workbook
    if "." in file_name:
        file_name = file_name.split(".")[0]+"_fmt.xlsx"
    elif "." not in file_name:
        file_name = file_name+"_fmt.xlsx"
    wb.save(filename = path+file_name)
    print("Finished formatting, saved.")
    
