using DocumentFormat.OpenXml.Packaging;
using DocumentMergeEngine;
using DocumentMergeEngine.Interfaces;
using OpenXMLTools;
using OpenXMLTools.Interfaces;
using System;
using System.IO;
using System.Windows;
using Service = DocumentTransitionApp.TransitionAppService;
using System.Collections.Generic;

namespace DocumentTransitionApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
	{
		Service.PersonFiles[] result;
        IPresentationTools PresentationTools;

		public MainWindow()
		{
            PresentationTools = new PresentationTools();
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
				ILocalMerge merge = new WordMerge();
				merge.Run(filename);
			}
			//Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			//var resultFile = serviceClient.MergeDocument(result);
			//int dupa = 1;
		}

        private void button_Click_4(object sender, RoutedEventArgs e)
        {
            MemoryStream inMemoryCopy = new MemoryStream();
            using (FileStream fs = File.OpenRead(@"C:\Users\drabiu\Documents\Testy\przykladowa-prezentacja.pptx"))
            {
                fs.CopyTo(inMemoryCopy);
            }

            MemoryStream inMemoryCopy2 = new MemoryStream();
            using (FileStream fs = File.OpenRead(@"C:\Users\drabiu\Documents\Testy\6.CGW15-prezentacja.pptx"))
            {
                fs.CopyTo(inMemoryCopy2);
            }

            byte[] byteArray = StreamTools.ReadFully(inMemoryCopy);
            byte[] byteArray2 = StreamTools.ReadFully(inMemoryCopy2);
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (PresentationDocument preDoc =
                   PresentationDocument.Open(mem, true))
                {                    
                    using (MemoryStream mem2 = new MemoryStream())
                    {
                        mem2.Write(byteArray2, 0, (int)byteArray2.Length);
                        using (PresentationDocument pre2Doc =
                   PresentationDocument.Open(mem2, true))
                        {                          
                            PresentationTools.InsertSlidesFromTemplate(preDoc, pre2Doc, new List<string>() { "rId13" });
                        }
                    }             
                }

                byteArray = mem.ToArray();
                //System.IO.File.WriteAllBytes(@"C:\Users\drabiu\Documents\Testy\przykladowa-prezentacja-test.pptx", mem.ToArray());
            }
            System.IO.File.WriteAllBytes(@"C:\Users\drabiu\Documents\Testy\przykladowa-prezentacja-test.pptx", byteArray);
        //System.IO.File.WriteAllBytes(@"C:\Users\drabiu\Documents\Testy\przykladowa-prezentacja-test.pptx", inMemoryCopy.ToArray());

            //using (FileStream file = new FileStream(@"C:\Users\drabiu\Documents\Testy\przykladowa-prezentacja.pptx", FileMode.Create, FileAccess.Write))
            //{
            //    inMemoryCopy.WriteTo(file);
            //}

            inMemoryCopy.Close();
        }    
    }
}
