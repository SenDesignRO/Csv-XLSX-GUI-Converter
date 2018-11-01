#pragma once

namespace CsvtoXLSXGUI {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Summary for MainWindow
	/// </summary>
	public ref class MainWindow : public System::Windows::Forms::Form
	{
	public:
		MainWindow(void)
		{
			InitializeComponent();
			// The config should be read at the start of the program.
			manage_config();
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~MainWindow()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  about;
	protected:
	private: System::Windows::Forms::Button^  browse_button1;
	private: void manage_config();

	private: System::Windows::Forms::Label^  destinationLabel;
	private: System::Windows::Forms::TextBox^  output_path;
	private: System::Windows::Forms::Button^  browse_button2;
	private: System::Windows::Forms::Button^  startButton;
	private: System::Windows::Forms::ListView^  fileList;
	private: System::Windows::Forms::ColumnHeader^  FileName;
	private: System::Windows::Forms::ColumnHeader^  FilePath;
	private: System::Windows::Forms::ColumnHeader^  FileStatus;
	private: System::Windows::Forms::Button^  deleteButton;




	protected:

	protected:

	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			this->browse_button1 = (gcnew System::Windows::Forms::Button());
			this->destinationLabel = (gcnew System::Windows::Forms::Label());
			this->output_path = (gcnew System::Windows::Forms::TextBox());
			this->browse_button2 = (gcnew System::Windows::Forms::Button());
			this->startButton = (gcnew System::Windows::Forms::Button());
			this->fileList = (gcnew System::Windows::Forms::ListView());
			this->FileName = (gcnew System::Windows::Forms::ColumnHeader());
			this->FilePath = (gcnew System::Windows::Forms::ColumnHeader());
			this->FileStatus = (gcnew System::Windows::Forms::ColumnHeader());
			this->deleteButton = (gcnew System::Windows::Forms::Button());
			this->about = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			// 
			// browse_button1
			// 
			this->browse_button1->Cursor = System::Windows::Forms::Cursors::Hand;
			this->browse_button1->Location = System::Drawing::Point(12, 13);
			this->browse_button1->Name = L"browse_button1";
			this->browse_button1->Size = System::Drawing::Size(75, 23);
			this->browse_button1->TabIndex = 0;
			this->browse_button1->Text = L"Add Files";
			this->browse_button1->UseVisualStyleBackColor = true;
			this->browse_button1->Click += gcnew System::EventHandler(this, &MainWindow::browse_button1_Click);
			// 
			// destinationLabel
			// 
			this->destinationLabel->AutoSize = true;
			this->destinationLabel->Location = System::Drawing::Point(93, 18);
			this->destinationLabel->Name = L"destinationLabel";
			this->destinationLabel->Size = System::Drawing::Size(63, 13);
			this->destinationLabel->TabIndex = 5;
			this->destinationLabel->Text = L"Destination:";
			// 
			// output_path
			// 
			this->output_path->Location = System::Drawing::Point(162, 15);
			this->output_path->Name = L"output_path";
			this->output_path->Size = System::Drawing::Size(455, 20);
			this->output_path->TabIndex = 4;
			// 
			// browse_button2
			// 
			this->browse_button2->Cursor = System::Windows::Forms::Cursors::Hand;
			this->browse_button2->Location = System::Drawing::Point(623, 15);
			this->browse_button2->Name = L"browse_button2";
			this->browse_button2->Size = System::Drawing::Size(75, 23);
			this->browse_button2->TabIndex = 3;
			this->browse_button2->Text = L"Browse";
			this->browse_button2->UseVisualStyleBackColor = true;
			this->browse_button2->Click += gcnew System::EventHandler(this, &MainWindow::browse_button2_Click);
			// 
			// startButton
			// 
			this->startButton->Cursor = System::Windows::Forms::Cursors::Hand;
			this->startButton->Location = System::Drawing::Point(704, 15);
			this->startButton->Name = L"startButton";
			this->startButton->Size = System::Drawing::Size(75, 23);
			this->startButton->TabIndex = 6;
			this->startButton->Text = L"Start";
			this->startButton->UseVisualStyleBackColor = true;
			this->startButton->Click += gcnew System::EventHandler(this, &MainWindow::startButton_Click);
			// 
			// fileList
			// 
			this->fileList->AllowDrop = true;
			this->fileList->BorderStyle = System::Windows::Forms::BorderStyle::FixedSingle;
			this->fileList->Columns->AddRange(gcnew cli::array< System::Windows::Forms::ColumnHeader^  >(3) {
				this->FileName, this->FilePath,
					this->FileStatus
			});
			this->fileList->Cursor = System::Windows::Forms::Cursors::Hand;
			this->fileList->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->fileList->Location = System::Drawing::Point(12, 44);
			this->fileList->Name = L"fileList";
			this->fileList->Size = System::Drawing::Size(920, 530);
			this->fileList->TabIndex = 7;
			this->fileList->UseCompatibleStateImageBehavior = false;
			this->fileList->View = System::Windows::Forms::View::Details;
			this->fileList->DragDrop += gcnew System::Windows::Forms::DragEventHandler(this, &MainWindow::DragDropHandler);
			this->fileList->DragOver += gcnew System::Windows::Forms::DragEventHandler(this, &MainWindow::DragOverHandler);
			this->fileList->KeyDown += gcnew System::Windows::Forms::KeyEventHandler(this, &MainWindow::fileList_KeyDown);
			// 
			// FileName
			// 
			this->FileName->Text = L"File Name";
			this->FileName->Width = 170;
			// 
			// FilePath
			// 
			this->FilePath->Text = L"Path";
			this->FilePath->Width = 686;
			// 
			// FileStatus
			// 
			this->FileStatus->Text = L"Status";
			// 
			// deleteButton
			// 
			this->deleteButton->Cursor = System::Windows::Forms::Cursors::Hand;
			this->deleteButton->Location = System::Drawing::Point(785, 15);
			this->deleteButton->Name = L"deleteButton";
			this->deleteButton->Size = System::Drawing::Size(75, 23);
			this->deleteButton->TabIndex = 8;
			this->deleteButton->Text = L"Delete";
			this->deleteButton->UseVisualStyleBackColor = true;
			this->deleteButton->Click += gcnew System::EventHandler(this, &MainWindow::deleteButton_Click_1);
			// 
			// about
			// 
			this->about->Cursor = System::Windows::Forms::Cursors::Hand;
			this->about->Location = System::Drawing::Point(866, 15);
			this->about->Name = L"about";
			this->about->Size = System::Drawing::Size(66, 23);
			this->about->TabIndex = 9;
			this->about->Text = L"About";
			this->about->UseVisualStyleBackColor = true;
			this->about->Click += gcnew System::EventHandler(this, &MainWindow::about_Click);
			// 
			// MainWindow
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(944, 586);
			this->Controls->Add(this->about);
			this->Controls->Add(this->deleteButton);
			this->Controls->Add(this->fileList);
			this->Controls->Add(this->startButton);
			this->Controls->Add(this->destinationLabel);
			this->Controls->Add(this->output_path);
			this->Controls->Add(this->browse_button2);
			this->Controls->Add(this->browse_button1);
			this->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedSingle;
			this->MaximizeBox = false;
			this->Name = L"MainWindow";
			this->ShowIcon = false;
			this->ShowInTaskbar = false;
			this->Text = L"CSV / XLSX Convertor";
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion

	private: System::Void browse_button1_Click(System::Object^  sender, System::EventArgs^  e);
	private: System::Void deleteButton_Click_1(System::Object^  sender, System::EventArgs^  e);
	private: System::Void browse_button2_Click(System::Object^  sender, System::EventArgs^  e);
	private: System::Void startButton_Click(System::Object^  sender, System::EventArgs^  e);
	private: System::Void DragDropHandler(Object ^sender, DragEventArgs ^args);
	private: System::Void DragOverHandler(Object ^sender, DragEventArgs ^args);
	private: System::Void fileList_KeyDown(System::Object^  sender, System::Windows::Forms::KeyEventArgs^  e);
	private: System::Void about_Click(System::Object^  sender, System::EventArgs^  e);
	
	};
}
