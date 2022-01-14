# copy_paste_automation
I tried to automate the process of copy-pasting data from one Excel workbook to another, with different formats. 

The painpoint of manual work in Excel is that it is easy to make trivial mistakes when dealing with middle-sized Excel datasets, which would cause unnecessary trouble on data quality.
So I decided to use Python to automate this process, leaving no space for the 'manual work devil' to take my place.  

Following the existed solutions, I successfully run the code at the beginning based on the instruction on: https://yagisanatode.com/2017/11/18/copy-and-paste-ranges-in-excel-with-openpyxl-and-python-3/   
However, it performed extremely slow on my computer - with Visual Studio, I directly chose the language as 'Python' and started writing scripts. I would need approximately 120 seconds to copy 6 cells in a workbook.  


I then started to decrease the time. At first I thought it was solely the problem of openpyxl. So I searched for new libraries such as xlwt or xlwings. However, each of the libraries could only realize the copy-pasting of a single row/column. A selective range of cells cannot be realized.  

