# Csv/XLSX-GUI-Converter
A simple CSV/XLSX convertor with a friendly GUI.

![Alt text](screenshot.jpg?raw=true)

This software was creaded using C++/CLI. It can parse CSV files delimited by comma (,) or quoted comma (","), then convert them to Excel files or XLSX Excel files and convert them into CSV files. It also takes in account escaped quotes and removes the extra ones. Drag and Drop is supported. A config file will be created at the first run and it is used to store the output location for futher uses and convenience. 

This program is based on the <b>libxlsxwriter</b> C++ library which can be found here https://github.com/jmcnamara/libxlsxwriter and the XLSX I/O C library which can be found here https://github.com/brechtsanders/xlsxio. The config parser is taken from <b>arcomber</b> on Stack, the question can be viewed <a href="https://codereview.stackexchange.com/questions/127819/ini-file-parser-in-c/127863">here</a>.

In order to run this program there are three dll files required, LibXlsxWriter.dll, libxlsxio_read.dll and Zlib.dll present in the same folder as the compiled exe. You may compile your own version of the dlls if you wish so from the code provided in the original repositories. 

You are free to use the final product or reuse the code in any way you like as long as you credit <b>me</b>, the <b>XLSX I/O</b> author and the <b>libxlsxwriter</b> author. 
