The Python script attached is capable of importing Data from Plain text file to Excel automatically. I use this to import data from ANSYS FLUENT runs into excel. You might need to change some of the code to make it suitable for use with your data file.

When you run this script from a folder it will only import data from files that ends with '.out'. If you want to import from files that have a different ending (e.g. .txt, .csv, .jsn, etc) you need to change line 26 of the code.

Also this script will only get the data present only in the first two columns of each csv file.
