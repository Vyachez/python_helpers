# Library of Python helper functions

This library of helper functions that can be used for Python projects related to data parsing, cleansing, preparation or any automation efforts.

## Getting Started

Explore **List of helpers** section below to find necessary helper function for you project.

### Prerequisites and Installation

Some functions may require additional modules and libraries to be installed.
Most common:
- pandas
- numpy

Use `pip install` to acquire necessary modules
Additional modules to be indicated for each helper function

## List of helpers
**_rename_files.py_** - Rename multiple files within directory

**_read_outlook.py_** - Extracting data from OutLook emails according to provided query.

_Additional modules requires to install:_
- pypiwin32
- zipfile36
- tqdm

_Example  of usage:_

`Mailbox(path=os.getcwd(), mailbox="yourlastname", folder=1, subj_keys=["Python User"],
	text=True, attach=True, unzip=True).search_mail()`
  
 _Instructions:_
  Please read class description or call `help(Mailbox)` command
  
### Contributing to this repository
- anyone can contribute to this repository via pull request
- new helper function should have short description to follow
- code within helper function should be well commented
- in case if helper function is complex it should have more thorough description and contain example of usage as above.

### License
Anyone can use this repository code for their own purpose and at own risk. There is no guarantee the code will work inside specific projects.
In case if advice, consultation required to embed the code into specific project, please send request to author.

### Authors
* **Viacheslav Nesterov** - Initial work

### Contributors
Please lis your name here if you contribute to this repository and indicate function/helper name you are contributing.

Viacheslav Nesterov - _read_outlook.py_

Viacheslav Nesterov - _rename_files.py_
