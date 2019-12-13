# CSV2VCF
Python Script for converting .csv contact files to .vcf. 

# xlsx2VCF
Python Script for converting .xlsx contact files to .vcf. 

# xlsx2csv utf8
Python Script for converting .xlsx contact files to .csv utf8 

This script is handy if you need to transfer contacts from a phone running windows to a new android/ iOS device. 

# IMPORTANT
If you find a bug or want to contribute new features, feel free to create a pull request.

Currently the following information will be converted:
- First Name 
- Last Name
- Home Phone
- Business Phone
- Home Fax
- Business Fax
- Home Address
- Business Address
- E-mail Address


# Usage
- make sure you have Python3 installed
- install using pip the following modules:
  - csv
  - unicodecsv
  - codecs
  - tkinter
  - collections
  - openpyxl
  - xlrd
  - argparse
  - webbrowser
  - cx_freeze
  - packaging
  - setuptools
  
- clone this repo
- using cmd to go to project folder using cd 
- write the following 
   ``` 
    python setup.py build 
    ```
- in folder build you will find file Excel to VCF.exe, double click on this file
then Gui of program will appear
- you will find 3 options:
  - Excel to VCF :  to convert contacts from xlsx to VCF file to be used for your phone
  - Excel to CSV utf-8 :  convert contacts from xlsx to CSV utf-8 file 
  - CSV to VCF : convert contacts from CSV utf-8 to VCF file to be used for your phone
