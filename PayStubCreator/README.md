# PayStubCreator

This app is a small experiment with using NuGet packages, namely [EPPlus](https://github.com/JanKallman/EPPlus) for reading excel files and [PDFsharp](http://www.pdfsharp.net/MainPage.ashx) for creating .pdf files.
The aim was to parse employee working hours in a set excel format (excel template available as paystub_template.xlsx) and calculate total pay, retirement and healthcare deductions as well as income tax and generate an appropriate .pdf pay stub.

## User interface

The user interface is created in WPF. It allows for the excel file selection, batch mode where all .xls and .xlsx files are parsed and optional company logo image selection.
![alt text](http://i.imgur.com/nzs4wl5.png "User interface")
The application has rudimentary logging built-in, which reports files where errors were encountered when trying to parse data.

## Output
Sample output pdf:

![Output pdf](https://i.imgur.com/KyNvZDV.png)

### Dependencies

- [EPPlus](https://github.com/JanKallman/EPPlus)
- [PDFsharp](http://www.pdfsharp.net/MainPage.ashx)

### Author's notes

Encountered exceptions are handled "unnaturally" via raising events and passing appropriate error messages to the LogError method. This was done as an extensive events and EventArgs practice.

License
----

MIT

**Free Software, Hell Yeah!**