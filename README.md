# excelReader
This is a simple class I wrote to for research. The code here currently reads the first column of an Excel file of numbers and turns it into a series of StartEnd objects.

StartEnd
: An object with two fields, an int start and end which determine a range of numbers between start and end.

# Developer Insights
I am currently using the Apacho POI in order to read the Excel file. Note that this code can easily be expanded to include multiple rows as well, if you wish to read tables of data besides just columns of data. The code can easily be edited to handle such cases.

There might be plans in the future to expand upon this so that it reads Strings that are a combination of numbers and characters, depending on the necessity of my research.

If you have any questions or issues, feel free to send me an email at rayf1013@cs.washington.edu.
