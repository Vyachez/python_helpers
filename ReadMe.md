# Library of Python helper functions

This library of helper functions that can be used for Python projects related to data parsing, cleansing, preparation or any automation efforts.

## Getting Started

Explore **List of helpers** section below to find necessary helper function for you project.

### Prerequisites and Installation

Some functions may require additional modules and libraries to be installed.
Most common:
- pandas
- numpy
- openpyxl
- tqdm

Use `pip install <module name>` to acquire necessary modules.
Additional modules to be indicated for each helper function

## List of helpers
**1. _rename_files.py_** - Rename multiple files within directory

**2. _read_outlook.py_** - Extracting data from OutLook emails according to provided query.

>_Additional modules required to install:_
- pypiwin32
- zipfile36

>_Example  of usage:_
`Mailbox(path=os.getcwd(), mailbox="yourlastname", folder=1, subj_keys=["Python User"],
	text=True, attach=True, unzip=True).search_mail()`
  
 >_Instructions:_
 Copy file into your working code directory. Import module into your code using `import read_outlook`. Read module class description or call `help(read_outlook.Mailbox)` command for all arguments documentation.
 
 **3. _magic_xl.py_** - Magic Excel. Automatic Excel Spreadsheet formatting.
>Additional modules required to install:
- string
>_Example  of usage:_
- `decor(path, file_name, tab_name)` - if you target certain Excel file in directory. `tab_name` is name of desired tab when formatted file returned. The file will be created in same directory.
- `decor(path, file_name, tab_name, dataframe=[])` - if you pass Pandas Dataframe to be formatted. The file will be created and saved in specified directory under `path`and will have `file_name`_fmt.xlsx name.
- `decor(path, "Grocery_shop.xlsx", tab_name='Fruits',
             dataframe=master_df,
             row_h=None, col_ws=[11, 17, 165],
             header_h=27, conditional_coloring={'apple':'34eb5f', 'banana':'e8eb34'},
             exact_condition=True,
         header_fill_color="DDDDDD", header_text_color='000000', body_color="FFFFFF",
         left_col_color="FFFFFF",
         footnote="open 24/7")` - full example with passing dataframe, including conditional formatting and specified columns width.
>Instructions and additional description:
Copy file into your working code directory. Import module into your code using `import magic_xl`. Read module class description or call `help(magic_xl.decor)` command for all arguments documentation.
  
### Contributing to this repository
- anyone can contribute to this repository with pull request
- new helper function should have short description and arguments documentation
- code within helper function should be well commented
- in case if helper function is complex it should have more thorough description and contain example of usage as above.

### License
Anyone can use this repository code for their own purpose and at own risk. There is no guarantee the code will work inside specific projects.
In case if advice, consultation required to embed the code into specific project, please send request to developer.

### Developers
Please list your name here if you contribute to this repository and indicate function/helper name you are contributing to.
>**Viacheslav Nesterov:**
> - _read_outlook.py_
> - _rename_files.py_
> - _magic_xl.py_