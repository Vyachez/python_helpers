# -*- coding: utf-8 -*-
"""
utlities file for data processing

"""
import pandas as pd
import numpy as np
import csv
from os import sys
import os
from tqdm import tqdm
import matplotlib.pyplot as plt
import chardet
import collections

class Util:
    '''Contains key functions to open data files as .csv and .xlsx'''
    def __init__(self, path=os.getcwd()):
        # root path - if not specified - gets everything from current dir
        self.path = path
        self.data_path = self.path + '/rawdat/'
        self.scheme_path = self.path + '/scheme/'
        self.output_path = self.path + '/csv_output/'
        
    def flist(self, path):
        '''getting list of files in path folder'''
        print('Directory: {}'.format(path))
        file_lst = []
        for path, dirs, files in os.walk(path):
            for file in files:
                file_lst.append(file)
        return file_lst
    
    def pd_csv_opener(self, path, pd_encoding, header):
        '''reading datafile (mostly used for columns processing)'''
        return pd.read_csv(path, 
                           encoding=pd_encoding,
                           dtype=object,
                           header=header)
        
    def pd_excel_reader(self, path, f_name, extension, sheet_name, header, dtype=object):
        '''reading datafile in excel format'''
        dat = pd.read_excel(path + f_name + extension, sheet_name=sheet_name, header=header, dtype=dtype)
        # if excel had tabs - the output will be dictionary of data frames so we need to concat them all        
        if type(dat) == collections.OrderedDict:
            data_list = []
            shapes = 0
            for d in dat.keys():
                shapes+=dat[d].shape[0]
                data_list.append(dat[d])
                data = pd.concat(data_list)
            assert data.shape[0] == shapes
        return data
    
    def test_csv_open(self, path, f_name, n_rows=100, header=0):
        '''open pregenerated csv to check output - limited number of rows'''
        return pd.read_csv(path + f_name + '.csv', 
                           encoding='latin-1',
                           dtype=object,
                           nrows=n_rows,
                           header=header)
        
    def find_encoding(self, fname):
        '''Detecting encoding for specified file'''
        print('Detecting encoding for',self.data_path+fname+'.txt')
        r_file = open(self.data_path+fname+'.txt', 'rb').read()
        result = chardet.detect(r_file)
        charenc = result['encoding']
        return charenc
    
    def alt_opener(self, path, name, extension, sheet_name, usecols, header, py_encoding, pd_encoding):
        '''file opener
        '''            
        # if file is more than 1 gig - will open it with chunks
        print('\nCurrent working directory: {}'.format(path))
        size = os.path.getsize(path+name+extension)/1000**3
        if size > 1:
            # getting extension to process with
            if extension == '.xlsx' or extension == '.xls':
                print("File is too big: {} Gb, to be processed by Excel reader in Pandas.".format(size) +\
                      "\nBut you can still try it. If fails, another option should be coded to approach the case.")
                rslt = input("\nDo you still want to try processing it? y/n")
                if str(rslt).lower() == 'y':
                    print("Ok, Processing...")
                    return self.pd_excel_reader(path, name, extension, sheet_name, header, dtype=object)
                else:
                    print("Ok, your choince was 'No'. Terminating...")
                    sys.exit()
            elif extension == '.txt' or extension == '.csv':
                print("File is too big: {} Gb, will be processed by chunks".format(size))
                # returning empty dataframe
                return pd.DataFrame()
            else:
                print("Extension specified '{}' is not recognized. Terminating.".format(extension))
        else:
            # getting extension to process with
            if extension == '.xlsx' or extension == '.xls':
                print("Opening Excel file...")
                return self.pd_excel_reader(path, name, extension, sheet_name, header, dtype=object)
            elif extension == '.txt' or extension == '.csv':
                print("Opening comma delimited file...")
                print('Encoding:',py_encoding)
                try:
                    with open(path + name + extension, 'rt', encoding=py_encoding, errors='replace', buffering=16*1000**2) as csvfile: # with 'replace' unrecognized symbols with '?'
                        dset = csv.reader(x.replace('\0', '') for x in csvfile) #removing null byte (zero column vals)
                        data = []
                        for row in dset:
                            data.append(row)            
                    csvfile.close()
                    data = pd.DataFrame(data)            
                    return data
                except:
                    print("Error occured while opening. Trying different method")
                    if usecols == None:
                        print("This type of file can't be processed without scheme. Please specify scheme file.")
                        sys.exit()
                    print('Encoding:',pd_encoding)
                    data = pd.read_csv(path + name + extension, encoding=pd_encoding, engine=None,
                                         converters=self.converter_dict(usecols),
                                         delimiter=',', quotechar='"', low_memory = False,
                                         error_bad_lines=False, warn_bad_lines = True,
                                         chunksize=self.chunksize, header=None, index_col=False,
                                         memory_map=True)
                    return data
            else:
                print("Extension specified '{}' is not recognized. Terminating.".format(extension))
    
    def converter_dict(self, names):
        '''prepares conversion dictionary to open files just with pandas read_csv
        most of files are being cracked easily with this converter
        see pplication in processing.py module'''
        # building delimiting function
        def delimit(char):
            if char != None:
                char = char.replace('"*','')
            return char
        
        # creating dictionary with function for each column to be parsed
        conv = {}
        for c in names:
            conv[c] = delimit
        return conv                 
        
    def header_maker(self, sch):
        '''making header for file'''
        '''file specific function'''
        header = {}
        for k, v in zip(sch[1:2].columns, sch.iloc[1][1:]):
            header[k] = v
        return header  
    
    def cols_maker(self, sch):
        '''making header for file'''
        '''file specific function'''
        cols = []
        for c in sch.iloc[1][1:]:
            cols.append(c)
        return cols
    
    def makedata(self, name):
        '''put together header and file in txt'''
        # reading datafile
        data = self.pd_csv_opener(self.data_path + name + '.txt')
        scheme = self.pd_csv_opener(self.scheme_path + name + '.csv')
        data = data.rename(columns=self.header_maker(scheme))
        return data
     
    # plotting function
    def bar_plot(self, x, y):
        '''plots dataset:
            needs x, y as list'''
        plt.figure(figsize=(10,5))
        plt.bar(np.array(x), np.array(y))
        plt.xticks(np.array(x), np.array(y))
        plt.show()
        
    def scatter_plot(self, x, y, c, s):
        fig, ax = plt.subplots()
        ax.scatter(x, y, c, s)
    #    ax.set_xlabel(r'$\Delta_i$', fontsize=15)
    #    ax.set_ylabel(r'$\Delta_{i+1}$', fontsize=15)
    #    ax.set_title('Volume and percent change')
        ax.grid(True)
        fig.tight_layout()
    # object setup
    #dtype={0: 'object', 1: 'object', 2: 'object', 3: 'object', 4: 'float64'}

class Cure:
    '''This class cures data based on columns provided and index.
    takes index and dataset as an argument'''
    def __init__(self, dset):
        self.dset = dset

    def one_col_shift(self, idx, col_a, col_b, delimiter, drop_q=False):
        '''shifts one column to left merging two neighbours
        clo_a - column to correct data in
        col_b - column that contains part of data for col_a
        and will be delted after correction for col_a
        delimiter - char to be inserterd in between two strings from columns
        drop_q - delete double quotation mark from string'''
        cols = self.dset.columns
        self.dset.loc[idx, col_a] = \
        self.dset.loc[idx, col_a] + delimiter + self.dset.loc[idx, col_b]
        if drop_q:
            if '"' in self.dset.loc[idx, col_a]:
                self.dset.loc[idx, col_a] = self.dset.loc[idx, col_a].replace('"', '')
        tmp_row = self.dset.loc[idx].drop(col_b)
        tmp_row = np.append(tmp_row,'')
        tmp_row = np.array([tmp_row])
        if len(tmp_row[0]) != len(cols):
            print("Error: row length {} does not match with header length {}!".format(len(tmp_row[0]),
                  len(cols)))
            os.exit()
        new_row = pd.DataFrame(tmp_row, columns=cols, index=None)
        self.dset.loc[idx] = new_row.values
        self.dset = self.dset.reset_index(drop=True)
        return self.dset
    
    def one_col_merge(self, idx, col_a, col_b, delimiter, drop_q=False):
        '''cutting value from column and adds it to left neighbour column
        clo_a - column to append data to
        col_b - column that contains data for col_a to cut and cleanup after
        delimiter - char to be inserterd in between two strings from columns
        drop_q - delete quotation mark from string'''
        self.dset.loc[idx, col_a] = \
        self.dset.loc[idx, col_a] + delimiter + self.dset.loc[idx, col_b]
        if drop_q:
            if '"' in self.dset.loc[idx, col_a]:
                self.dset.loc[idx, col_a] = self.dset.loc[idx, col_a].replace('"', '')
        self.dset.loc[idx, col_b] = ''
        return self.dset
    
    def split_to_right(self, cols, idx, bcol, steps, delimiter, drops=None):
        '''split breakdown column with bad delimiter and shifts only selected columns 
        (not all row) to right
        cols - list of columns to fix (including breakdown)
        idx - index to fix
        bcol - breakdown column
        steps - number of columns to shift
        delimiter - symbols to split on
        drops - list of symbols to cleanup from field'''
        
        # gettin goriginal columns
        orig_cols = self.dset.columns
        
        # making new dataset with single row to correct
        tmp_set = pd.DataFrame(self.dset.loc[self.dset.index == idx], columns=orig_cols)
        
        # getting content in column with delimiter and amending it
        part_1 = "".join(tmp_set.loc[idx, bcol].split(delimiter)[:-1])
        part_2 = "".join(tmp_set.loc[idx, bcol].split(delimiter)[1:])
        if drops != None:
            for d in drops:
                try:
                    part_1 = part_1.replace(d, '')
                except:
                    pass
                try:
                    part_2 = part_2.replace(d, '')
                except:
                    pass
                          
        tmp_set.loc[idx, bcol] = part_1
        
        # shifting columns right (adding empty) and adding second part of breakdown column
        emp_cols = ['temp_col'] * steps
        col_idx = tmp_set.columns.get_loc(bcol)
        for c, i in zip(emp_cols, range(steps)):
            tmp_set.insert(col_idx+i+1, c, '',  allow_duplicates=True)
        
        # deleting leftovers
        tmp_set = tmp_set.drop(columns=tmp_set.columns[-steps:])
        
        # updating original columns
        tmp_set.columns = orig_cols
        
        # updating value in second part of breakdown column
        tmp_set.loc[idx, tmp_set.columns[col_idx+steps]] = part_2
        
        # updating only columns to be fixed
        for c in cols:
            # replacing row in original dataset
            self.dset.loc[ self.dset.index == idx, c] = tmp_set.loc[tmp_set.index == idx, c]
        
        # resetting index
        self.dset =  self.dset.reset_index(drop=True)
        return  self.dset
    
    def share_to_left(self, idx, bcol, lcol, delimiter, drops=None):
        '''split breakdown column with bad delimiter and share one part of it to left
        idx - index to fix - list
        bcol - breakdown column
        lcol - column to repair on left
        delimiter - symbols to split on
        drops - list of symbols to cleanup from field'''
        
        # gettin goriginal columns
        orig_cols = self.dset.columns
        
        print("Fixing each index element")
        for i in idx:
            # making new dataset with single row to correct
            tmp_set = pd.DataFrame(self.dset.loc[self.dset.index == i], columns=orig_cols)
            
            # getting content in column with delimiter and amending it
            part_1 = "".join(tmp_set.loc[i, bcol].split(delimiter)[:-1])
            part_2 = "".join(tmp_set.loc[i, bcol].split(delimiter)[1:])
            if drops != None:
                for d in drops:
                    try:
                        part_1 = part_1.replace(d, '')
                    except:
                        pass
                    try:
                        part_2 = part_2.replace(d, '')
                    except:
                        pass
                    
            tmp_set.loc[i, lcol] = tmp_set.loc[i, lcol] + " " + part_1                          
            tmp_set.loc[i, bcol] = part_2
            
            # replacing row in original dataset
            self.dset.loc[self.dset.index == i] = tmp_set.loc[tmp_set.index == i]
            
        # resetting index
        self.dset =  self.dset.reset_index(drop=True)
        return  self.dset
    
    def row_stretch(self, err_idx, delimiter, drops=None):
        '''STRETCHING VS NONES - AUTO:
        stretching row submitted within 'err_idx' list of indexes with breakdown column for each
        and shifting right for identified number of columns depending on identified NONES on right.
        rdxs - list of indexes to fix
        delimiter - symbols to split on (should be specified carefully 
          because it will help to identify and verify steps to shift - 
          specify only those which caused columns to clog)
        drops - list of symbols to cleanup from field'''
        
        def stretching(dset, idx, bcol, steps, delimiter, drops=None):
            '''stretching single row with breakdown column and shifting right for
            steps number of columns.
            idx - index to fix
            bcol - breakdown column
            steps - number of columns to shift
            delimiter - symbols to split on
            drops - list of symbols to cleanup from field'''
            
            # gettin goriginal columns
            orig_cols = dset.columns
            
            # making new dataset with single row to correct
            tmp_set = pd.DataFrame(dset.loc[dset.index == idx], columns=orig_cols)
            
            # getting content in column with delimiter and amending it
            part_1 = "".join(tmp_set.loc[idx, bcol].split(delimiter)[:-1])
            part_2 = "".join(tmp_set.loc[idx, bcol].split(delimiter)[1:])
            if drops != None:
                for d in drops:
                    try:
                        part_1 = part_1.replace(d, '')
                    except:
                        pass
                    try:
                        part_2 = part_2.replace(d, '')
                    except:
                        pass
                              
            tmp_set.loc[idx, bcol] = part_1
            
            # shifting columns right (adding empty) and adding second part of breakdown column
            emp_cols = ['temp_col'] * steps
            col_idx = tmp_set.columns.get_loc(bcol)
            for c, i in zip(emp_cols, range(steps)):
                tmp_set.insert(col_idx+i+1, c, '',  allow_duplicates=True)
            
            # deleting leftovers
            tmp_set = tmp_set.drop(columns=tmp_set.columns[-steps:])
            
            # updating original columns
            tmp_set.columns = orig_cols
            
            # updating value in second part of breakdown column
            tmp_set.loc[idx, tmp_set.columns[col_idx+steps]] = part_2
            
            # replacing row in original dataset
            dset.loc[dset.index == idx] = tmp_set.loc[tmp_set.index == idx]
            
            # resetting index
            dset = dset.reset_index(drop=True)
            return dset
        
        non_count = 0
        err_dict = {}
        seps_dict = {}
        
        if len(err_idx) > 0:
        
            # index preparation - finding Nones on rightmost columns to define how
            # many steps to shift
            print("index preparation")
            for e in err_idx:
                for i in reversed(self.dset.columns):
                    if self.dset.loc[e, i] == None:
                        non_count += 1
                    else:
                        break
                if non_count > 0:
                    err_dict[e] = non_count
                non_count = 0
            # finding columns that contain 'bad' delimiter
            print("finding columns that contain 'bad' delimiter")
            for k, v in zip(err_dict.keys(), err_dict.values()):
                for i in self.dset.columns:
                    try:
                        seps_count = self.dset.loc[k, i].count('",')
                        if seps_count > 0:
                            seps_dict[k] = i
                            #print('id: {}, separators: {}'.format(k, seps_count))
                    except:
                        pass 
                    
            #  verifying if the clog problem is only in one column otherwise will need to apply
            # different fix
            if len(seps_dict.keys()) != len(np.unique(list(seps_dict.keys()))):
                print("Row contains more than one columns with separators. Ambiguous operation. Use different fixing method")
                
            # now iterating through each index element and stretching
            print("stretching each index element")
            for s in seps_dict:
                #print("fixing idx:{}, columns {}".format(s, seps_dict.get(s)))
                self.dset = stretching(self.dset, s, seps_dict.get(s), err_dict.get(s), delimiter, drops=drops)
        
        else:
            print("Stretching fix is not needed")    
        
        return self.dset
    
    def stretch_by_idx(self, idx, bcol, steps, delimiter, drops=None):
        '''STRETCHING BASED ON Index and column parameters (not AUTO)
        stretching single row with breakdown column and shifting right for
        steps - defined number of columns.
        idx - index to fix
        bcol - breakdown column
        steps - number of columns to shift
        delimiter - symbols to split on
        drops - list of symbols to cleanup from field'''
        
        # getting original columns
        orig_cols = self.dset.columns
        
        # making new dataset with single row to correct
        tmp_set = pd.DataFrame(self.dset.loc[self.dset.index == idx], columns=orig_cols)
        
        # getting content in column with delimiter and amending it
        part_1 = "".join(tmp_set.loc[idx, bcol].split(delimiter)[:-1])
        part_2 = "".join(tmp_set.loc[idx, bcol].split(delimiter)[1:])
        if drops != None:
            for d in drops:
                try:
                    part_1 = part_1.replace(d, '')
                except:
                    pass
                try:
                    part_2 = part_2.replace(d, '')
                except:
                    pass
                          
        tmp_set.loc[idx, bcol] = part_1
        
        # shifting columns right (adding empty) and adding second part of breakdown column
        emp_cols = ['temp_col'] * steps
        col_idx = tmp_set.columns.get_loc(bcol)
        for c, i in zip(emp_cols, range(steps)):
            tmp_set.insert(col_idx+i+1, c, '',  allow_duplicates=True)
        
        # deleting leftovers
        tmp_set = tmp_set.drop(columns=tmp_set.columns[-steps:])
        
        # updating original columns
        tmp_set.columns = orig_cols
        
        # updating value in second part of breakdown column
        tmp_set.loc[idx, tmp_set.columns[col_idx+steps]] = part_2
        
        # replacing row in original dataset
        self.dset.loc[self.dset.index == idx] = tmp_set.loc[tmp_set.index == idx]
        
        # resetting index
        self.dset = self.dset.reset_index(drop=True)
        return self.dset

    
class Lookup:
    '''This class creates views, dataframes to preview data'''
    def __init__(self, dset):
        self.dset = dset

    def spar_index(self, idxs):
        '''builds new dataframe with selected indexes (listed) from dataset''' 
        cols = self.dset.columns
        newframe = pd.DataFrame(columns = cols)
        for i in idxs:
            newframe = newframe.append(self.dset.loc[i])
        return newframe

    def values(self, vals, col):
        '''builds new dataframe with selected values in selected columns
        vals as list
        cols as list''' 
        cols = self.dset.columns
        newframe = pd.DataFrame(columns = cols)
        for i in vals:
            newframe = newframe.append(self.dset.loc[self.dset[col] == i])
        return newframe

    def strings(self, vals, col):
        '''builds new dataframe with matched strings in values in selected columns
        vals as list
        cols as list''' 
        cols = self.dset.columns
        newframe = pd.DataFrame(columns = cols)
        for i in vals:
            for c in col:
                try:
                    newframe = newframe.append(self.dset.loc[self.dset[c].str.contains(i) == True])
                except:
                    print("Column {}, value {} parsing error".format(c,i))
                    raise
        return newframe

    def find_nones(self):    
        # finding Nones on rightmost columns to define
        non_count = 0
        err_idx = []
        err_dict = {}
        for c in self.dset.columns:
            if self.dset[c].isnull().any():
                err_idx += list(self.dset.loc[self.dset[c].isnull()].index)
        for e in list(set(err_idx)):
            for i in reversed(self.dset.columns):
                print(e, i)
                if self.dset.loc[e, i] == None:
                    non_count += 1
                else:
                    break
            if non_count > 0:
                err_dict[e] = non_count
            non_count = 0
        return err_dict
    
    def find_blanks(self, n_cols):
        'n_cols - number of columns threshold'
        err_idx = []
        print('Finding blanks')
        try:
            for i in tqdm(self.dset.index):
                if len(np.unique(self.spar_index([i]).fillna('')))<=n_cols:
                    err_idx.append(i)
        except:
            print("Error: did you pass number of threshold columns?")
            raise
        return err_idx
    