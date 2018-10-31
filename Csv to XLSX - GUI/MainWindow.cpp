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

string convertToString(String^ string) {
	using System::Runtime::InteropServices::Marshal;

	System::IntPtr pointer = Marshal::StringToHGlobalAnsi(string);
	char* charPointer = reinterpret_cast<char*>(pointer.ToPointer());
	std::string returnString(charPointer, string->Length);
	Marshal::FreeHGlobal(pointer);
	
	return returnString;
}

string returnFileName(String^ path) {
	string filename = convertToString(path);
	const size_t last_slash_idx = filename.find_last_of("\\/");
	if (std::string::npos != last_slash_idx)
	{
		filename.erase(0, last_slash_idx + 1);
	}
	return filename;
}

bool isCSV(string fn) {
	
	if (fn.substr(fn.find_last_of(".") + 1) == "csv") {
		return true;
	}
	return false;
}

bool fexists(const char *filename)
{
	ifstream ifile(filename);
	if (ifile) {
		return true;
	}
	return false;
}

void generate_config(const std::string& filename) {
	ofstream ostrm;
	ostrm.open(filename);
	if (ostrm) {
		ostrm << "[paths]\noutputPath = ";
	}
}

void MainWindow::manage_config()
{
	if (fexists("config.ini")) {
		config cfg("config.ini");
		section* pathsection = cfg.get_section("paths");

		if (pathsection != NULL) {
			String^ aux = gcnew String(cfg.get_value("paths", "outputPath").c_str());
			output_path->Text = aux;
		
		} else {
			generate_config("config.ini");
		}
	}
	else {
		generate_config("config.ini");
	}
}

void change_config(string section, string key, string value) {
	ofstream ostrm;
	ostrm.open("config.ini");
	if (ostrm) {
		ostrm << "["<<section<<"]"<<"\n"<<key<<" = "<<value;
	}
}

void write_worksheet_data(lxw_worksheet* worksheet, lxw_format* bold, string inputPath) {
	string line;
	int i = 1;
	char j = 'A';
	ifstream myfile(inputPath);
	if (myfile.is_open()) {
		while (getline(myfile, line)) {
			int type = 0;
			string delimiter;
			if (line.find("\",\"") != std::string::npos) {
				delimiter = "\",\"";
				type = 1;
				line.erase(0, 1);
				line.erase(line.size() - 1);
			}
			else {
				delimiter = ",";
				type = 0;
			}
			size_t pos = 0;

			while ((pos = line.find(delimiter)) != string::npos) {
				string substr = line.substr(0, pos);
				string x;
				x += j;
				x += to_string(i);
				const char* c = x.c_str();
				const char* b = substr.c_str();
				worksheet_write_string(worksheet, CELL(c), b, bold);
				j++;
				line.erase(0, pos + delimiter.length());
			}
			string x;
			x += j;
			x += to_string(i);
			line = regex_replace(line, regex("\"\""), "\"");
			const char* c = x.c_str();
			const char* b = line.c_str();
			worksheet_write_string(worksheet, CELL(c), b, bold);
			i++;
			j = 'A';
		}
		myfile.close();
	}
	else cout << "Unable to open file";
}

bool workbook(string inputPath, string outputPath, int item)
{
	const char * path = outputPath.c_str();
	lxw_workbook* workbook = new_workbook(path);
	lxw_worksheet* worksheet = workbook_add_worksheet(workbook, NULL);
	lxw_format* bold = workbook_add_format(workbook);
	write_worksheet_data(worksheet, bold, inputPath);
	return workbook_close(workbook);

}

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

Void MainWindow::startButton_Click(System::Object^  sender, System::EventArgs^  e) {
	if (!String::IsNullOrWhiteSpace(output_path->Text)) {
		for (int i = 0; i < fileList->Items->Count; i++)
		{
			if (!String::Equals(fileList->Items[i]->SubItems[2]->Text, L"Finished")) {
				string filename = convertToString(fileList->Items[i]->SubItems[0]->Text);
				const size_t period_idx = filename.rfind('.');
				if (std::string::npos != period_idx)
				{
					filename.erase(period_idx);
				}
				string outputPath = convertToString(output_path->Text) + filename + ".xlsx";
				if (workbook(convertToString(fileList->Items[i]->SubItems[1]->Text), outputPath, i) == 0) {
					fileList->Items[i]->SubItems[2]->Text = L"Finished";
				}
				else {
					fileList->Items[i]->SubItems[2]->Text = L"Error";
				}
			}
		}
	}
	else {
		MessageBox::Show("You must select an output folder before starting the conversion.");
	}
}

Void MainWindow::DragOverHandler(Object ^sender, DragEventArgs ^args) {

	if (args->Data->GetDataPresent(DataFormats::FileDrop))
		args->Effect = args->AllowedEffect & DragDropEffects::Link;
	else
		args->Effect = DragDropEffects::None;
}

Void MainWindow::DragDropHandler(Object ^sender, DragEventArgs ^args) {
	if (args->Data->GetDataPresent(DataFormats::FileDrop) && args->Effect == DragDropEffects::Link) {
		auto filePaths = dynamic_cast<cli::array<String^>^>(args->Data->GetData(DataFormats::FileDrop));

		if (filePaths != nullptr && filePaths->Length > 0) {
			int i = 0;
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