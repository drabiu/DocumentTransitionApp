using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;

using DocumentSplitEngine;
using DocumentMergeEngine;
using Service = DocumentTransitionApp.TransitionAppService;
using DocumentMergeEngine.Interfaces;

namespace DocumentTransitionApp
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		Service.PersonFiles[] result;

		public MainWindow()
		{
			InitializeComponent();
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			// Create OpenFileDialog 
			Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

			// Set filter for file extension and default file extension 
			dlg.DefaultExt = ".png";
			dlg.Filter = "Word Files (*.docx)|*.docx|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif";


			// Display OpenFileDialog by calling ShowDialog method 
			Nullable<bool> result = dlg.ShowDialog();

			// Get the selected file name and display in a TextBox 
			if (result == true)
			{
				// Open document 
				string filename = dlg.FileName;
				docxTextBox.Text = filename;
			}
		}

		private void Button_Click_1(object sender, RoutedEventArgs e)
		{
			
			// Create OpenFileDialog 
			Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

			// Set filter for file extension and default file extension 
			dlg.DefaultExt = ".png";
			dlg.Filter = "XML Files (*.xml)|*.xml|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif";


			// Display OpenFileDialog by calling ShowDialog method 
			Nullable<bool> result = dlg.ShowDialog();

			// Get the selected file name and display in a TextBox 
			if (result == true)
			{
				string filename = dlg.FileName;
				xmlTextBox.Text = filename;
			}
		}

		private void Button_Click_2(object sender, RoutedEventArgs e)
		{
			string docName = System.IO.Path.GetFileNameWithoutExtension(docxTextBox.Text);
			//ILocalSplit run = new DocumentSplit(docName);
			//run.OpenAndSearchWordDocument(docxTextBox.Text, xmlTextBox.Text);
			RunSplitWebService(docName, docxTextBox.Text, xmlTextBox.Text);
			//run.SaveSplitDocument(docxTextBox.Text);
		}

		private async void RunSplitWebService(string docName, string filePath, string xmlPath)
		{
			Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();

			byte[] fileStream = File.ReadAllBytes(filePath);
			byte[] xmlStream = File.ReadAllBytes(xmlPath);

			//result = serviceClient.SplitDocument(docName, fileStream, xmlStream);
			//Service.SplitDocumentResponse response = await serviceClient.SplitDocumentAsync(docName, fileStream, xmlStream);
			var result = serviceClient.GetParts(docName, fileStream);
		}

		private void Button_Click_3(object sender, RoutedEventArgs e)
		{
			Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();

			// Set filter for file extension and default file extension 
			dlg.DefaultExt = ".docx";
			dlg.Filter = "Word Files (*.docx)|*.docx|PNG Files (*.png)|*.png|JPG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif";


			//Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

			//Get the selected file name and display in a TextBox
            if (result == true)
			{
				// Open document 
				string filename = dlg.FileName;
				ILocalMerge merge = new DocumentMerge();
				merge.Run(filename);
			}
			//Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			//var resultFile = serviceClient.MergeDocument(result);
			//int dupa = 1;
		}
	}
}
