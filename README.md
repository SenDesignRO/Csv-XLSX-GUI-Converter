# Csv-to-XLSX-GUI-Converter
A simple CSV to XLSX convertor with a friendly GUI.

![Alt text](screenshot.jpg?raw=true)

This software was creaded using C++/CLI. It can parse CSV files delimited by comma (,) or quoted comma (","). It also takes in account escaped quotes and removes the extra ones. Drag and Drop is supported. A config file will be created at the first run and it is used to store the output location for futher uses and convenience. 

This program is based on the <b>libxlsxwriter</b> C++ library which can be found here https://github.com/jmcnamara/libxlsxwriter. The config parser is taken from <b>arcomber</b> on Stack, the question can be viewed <a href="https://codereview.stackexchange.com/questions/127819/ini-file-parser-in-c/127863">here</a>.

In order to run this program there are two dll files that must be present, LibXlsxWriter.dll and Zlib.dll present in the same folder as the compiled exe. You may compile your own version of the dlls if you wish so from the code provided in the original repository. 

You are free to use the final product or reuse the code in any way you like as long as you credit <b>me</b> and the <b>libxlsxwriter</b> author. 
