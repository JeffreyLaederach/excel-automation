# Overview

**This repository is for testing different ways of automating Microsoft Excel workbooks and their contents. The code for generating the Excel files can be found in the "Scripts" sub-folder in 3 different programming languages: Python, PowerShell, and Perl. The generated spreadsheets are located in the "Spreadsheets" sub-folder.**

## Python: 

In order to run the python file, you must have the XlsxWriter package installed in your Python environment which can be achieved in Anaconda using `conda install XlsxWriter`.

## PowerShell:

The PowerShell script can be executed via the Windows PowerShell Terminal or from the one integrated in Visual Studio Code by simply clicking the "Run Code" button on the top right.

## Perl:

1. Download the Strawberry Perl Installer "5.38.2.2 MSI (171.7 MB)" from the following site [Strawberry Perl Download](https://strawberryperl.com/)

2. Once it is installed, open up the Windows Command Prompt and type in `perl -v` to verify that Perl is properly accessible on your computer.

3. Install Excel-Writer-XLSX by typing the command `cpanm Excel::Writer::XLSX` into your command prompt terminal.

5. You can now create an Excel File using Perl by executing the script using the command structure found below:

`perl C:\Users\jeffl\OneDrive\Documents\GitHub\excel-automation\Scripts\CreateExcelFile.pl`
  
*Note: Be sure to appropriately edit the path of the script file in the code so that Perl can successfully locate the file to run.*
