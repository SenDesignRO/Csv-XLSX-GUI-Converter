#include "stdafx.h"
#include "MainWindow.h"
#include "xlsxwriter.h"
#include "config.h"

#include <string>
#include <iostream>
#include <fstream>
#include <regex>
#include <windows.h>

using namespace std;
using namespace System;
using namespace CsvtoXLSXGUI;

// Converts System::Strings to std::strings.
string convertToString(String^ string) {
	using System::Runtime::InteropServices::Marshal;

	System::IntPtr pointer = Marshal::StringToHGlobalAnsi(string);
	char* charPointer = reinterpret_cast<char*>(pointer.ToPointer());
	std::string returnString(charPointer, string->Length);
	Marshal::FreeHGlobal(pointer);

	return returnString;
}

// Returns the name of a file from a given path by identifying the last slash.
string returnFileName(String^ path) {
	string filename = convertToString(path);
	const size_t last_slash_idx = filename.find_last_of("\\/");
	if (std::string::npos != last_slash_idx)
	{
		filename.erase(0, last_slash_idx + 1);
	}
	return filename;
}

// Verifies if a given file is a CSV file.
bool isCSV(string filename) {

	if (filename.substr(filename.find_last_of(".") + 1) == "csv") {
		return true;
	}
	return false;
}

// Verifies if a file exists. Used to verify if the config file exists or needs to be created.
bool fexists(const char *filename)
{
	ifstream ifile(filename);
	if (ifile) {
		return true;
	}
	return false;
}

// Generates the config file.
void generate_config(const std::string& filename) {
	ofstream ostrm;
	ostrm.open(filename);
	if (ostrm) {
		ostrm << "[paths]\noutputPath = ";
	}
}

// Verifies if the config exists and reads it. If there is no config files, it creates one.
void MainWindow::manage_config()
{
	if (fexists("config.ini")) {
		config cfg("config.ini");
		section* pathsection = cfg.get_section("paths");

		if (pathsection != NULL) {
			String^ aux = gcnew String(cfg.get_value("paths", "outputPath").c_str());
			output_path->Text = aux;

		}
		else {
			generate_config("config.ini");
		}
	}
	else {
		generate_config("config.ini");
	}
}

// Changes the value of the default output path in the config file.
void change_config(string section, string key, string value) {
	ofstream ostrm;
	ostrm.open("config.ini");
	if (ostrm) {
		ostrm << "[" << section << "]" << "\n" << key << " = " << value;
	}
}

// Writes the given data in the given XLSX file.
void write_worksheet_data(lxw_worksheet* worksheet, lxw_format* bold, string inputPath) {
	string line;
	int i = 1;
	char j = 'A';
	// First it parses the CSV file.
	ifstream myfile(inputPath);
	if (myfile.is_open()) {
		while (getline(myfile, line)) {
			string delimiter;
			// It searches for the type of the delimiter. Either , or ",".
			if (line.find("\",\"") != std::string::npos) {
				delimiter = "\",\"";
				line.erase(0, 1);
				line.erase(line.size() - 1);
			}
			else {
				delimiter = ",";
			}
			size_t pos = 0;
			// Breaks each string into tokens starting from the delimiter and going backwards.
			while ((pos = line.find(delimiter)) != string::npos) {
				// By default, quotes are escaped. This removes the double quotes.
				line = regex_replace(line, regex("\"\""), "\"");
				string substr = line.substr(0, pos);
				string x;
				x += j;
				x += to_string(i);
				const char* c = x.c_str();
				const char* b = substr.c_str();
				// Writes the data to the Excel file.
				worksheet_write_string(worksheet, CELL(c), b, bold);
				j++;
				// Erases the read data from the buffer.
				line.erase(0, pos + delimiter.length());
			}
			// One last run to cover the last token after the last delimiter. It is required because the previous functions goes backwards and loses the final part.
			string x;
			x += j;
			x += to_string(i);
			// By default, quotes are escaped. This removes the double quotes.
			line = regex_replace(line, regex("\"\""), "\"");
			const char* c = x.c_str();
			const char* b = line.c_str();
			worksheet_write_string(worksheet, CELL(c), b, bold);
			i++;
			j = 'A';
		}
		myfile.close();
	}
}

// It creates the XLSX file at the given location and passes the CSV file to the parsing function. Returns 0 or an error if the file could not be created.
bool workbook(string inputPath, string outputPath, int item)
{
	const char * path = outputPath.c_str();
	lxw_workbook* workbook = new_workbook(path);
	lxw_worksheet* worksheet = workbook_add_worksheet(workbook, NULL);
	lxw_format* style = workbook_add_format(workbook);
	write_worksheet_data(worksheet, style, inputPath);
	return workbook_close(workbook);
}

// The function for the "Add" button. It adds the choosen file to the main list.
Void MainWindow::browse_button1_Click(System::Object^  sender, System::EventArgs^  e)
{
	OpenFileDialog^ openFileDialog1 = gcnew OpenFileDialog;

	openFileDialog1->InitialDirectory = "c:\\";
	openFileDialog1->Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
	openFileDialog1->FilterIndex = 2;
	openFileDialog1->RestoreDirectory = true;
	if (openFileDialog1->ShowDialog() == ::DialogResult::OK);
	{
		String^ folderName = openFileDialog1->FileName;
		ListViewItem^  listViewItem1 = (gcnew System::Windows::Forms::ListViewItem(gcnew cli::array< System::String^  >(3) {
			gcnew String(returnFileName(folderName).c_str()), folderName, L"Pending..."
		}, -1));
		this->fileList->Items->AddRange(gcnew cli::array< System::Windows::Forms::ListViewItem^  >(1) { listViewItem1 });

	}
}

// The function for the "Browse" button. It is used to point to the output path and add it to both the label and the config file.
Void MainWindow::browse_button2_Click(System::Object^  sender, System::EventArgs^  e)
{
	FolderBrowserDialog^ folderBrowserDialog = gcnew FolderBrowserDialog;
	if (folderBrowserDialog->ShowDialog() == ::DialogResult::OK);
	{
		String^ folderName = folderBrowserDialog->SelectedPath + "\\";
		output_path->Text = folderName;
		change_config("paths", "outputPath", convertToString(folderName));
	}
}

// The start button parses the list of files sending the file paths to the writer function. It also creates and sends the output path for each CSV file.
Void MainWindow::startButton_Click(System::Object^  sender, System::EventArgs^  e) {

	// Verifies if the Outpuh directory exists.
	if (!String::IsNullOrWhiteSpace(output_path->Text)) {
		for (int i = 0; i < fileList->Items->Count; i++)
		{
			// Skips the already converted files. 
			if (!String::Equals(fileList->Items[i]->SubItems[2]->Text, L"Finished")) {
				string filename = convertToString(fileList->Items[i]->SubItems[0]->Text);
				// Removes the extension from the filename. 
				const size_t period_idx = filename.rfind('.');
				if (std::string::npos != period_idx)
				{
					filename.erase(period_idx);
				}
				// Generates the output filenames from the input ones and adds the new extension. 
				string outputPath = convertToString(output_path->Text) + filename + ".xlsx";
				// If the file is sucessfuly created, it marks the conversion as finished in the third column of the list. 
				if (workbook(convertToString(fileList->Items[i]->SubItems[1]->Text), outputPath, i) == 0) {
					fileList->Items[i]->SubItems[2]->Text = L"Finished";
				}
				else {
					// Otherwise, it displays an error.
					fileList->Items[i]->SubItems[2]->Text = L"Error";
				}
			}
		}
	}
	else {
		MessageBox::Show("You must select an output folder before starting the conversion.");
	}
}

// Handles the drag over effect. It only allows files to be dropped into the list. 
Void MainWindow::DragOverHandler(Object ^sender, DragEventArgs ^args) {

	if (args->Data->GetDataPresent(DataFormats::FileDrop))
		args->Effect = args->AllowedEffect & DragDropEffects::Link;
	else
		args->Effect = DragDropEffects::None;
}

// Handles the Drag and Drop operation. 
Void MainWindow::DragDropHandler(Object ^sender, DragEventArgs ^args) {
	// Verifies if a file was dropped. 
	if (args->Data->GetDataPresent(DataFormats::FileDrop) && args->Effect == DragDropEffects::Link) {
		auto filePaths = dynamic_cast<cli::array<String^>^>(args->Data->GetData(DataFormats::FileDrop));
		// Verifies if the file is valid.
		if (filePaths != nullptr && filePaths->Length > 0) {
			int i = 0;
			// For each file dropped, it verifies if it's a CSV file and then it adds it to the list.
			for each (String^ path in filePaths) {
				if (isCSV(returnFileName(filePaths[i]).c_str())) {
					ListViewItem^  listViewItem1 = (gcnew System::Windows::Forms::ListViewItem(gcnew cli::array< System::String^  >(3) {
						gcnew String(returnFileName(filePaths[i]).c_str()), filePaths[i], L"Pending..."
					}, -1));
					this->fileList->Items->AddRange(gcnew cli::array< System::Windows::Forms::ListViewItem^  >(1) { listViewItem1 });
					i++;
				}
			}
		}
	}
}

// Deletes selected items from the list.
Void MainWindow::deleteButton_Click_1(System::Object^  sender, System::EventArgs^  e) {
	if (fileList->SelectedItems->Count > 0)
	{
		for (int i = fileList->Items->Count - 1; i >= 0; i--)
		{
			if (fileList->Items[i]->Selected)
			{
				fileList->Items[i]->Remove();
			}
		}
	}
}