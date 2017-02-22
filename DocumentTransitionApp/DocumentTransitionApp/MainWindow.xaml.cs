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
using OpenXMLTools.Interfaces;
using OpenXMLTools;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentSplitEngine.Data_Structures;

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
				ILocalMerge merge = new DocumentMerge();
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

            byte[] byteArray = ReadFully(inMemoryCopy);
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (PresentationDocument preDoc =
                   PresentationDocument.Open(mem, true))
                {
                    PresentationTools.InsertNewSlide(preDoc, 1, "aaaa");                 
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

        public static byte[] ReadFully(Stream stream)
        {
            long originalPosition = 0;

            if (stream.CanSeek)
            {
                originalPosition = stream.Position;
                stream.Position = 0;
            }

            try
            {
                byte[] readBuffer = new byte[4096];

                int totalBytesRead = 0;
                int bytesRead;

                while ((bytesRead = stream.Read(readBuffer, totalBytesRead, readBuffer.Length - totalBytesRead)) > 0)
                {
                    totalBytesRead += bytesRead;

                    if (totalBytesRead == readBuffer.Length)
                    {
                        int nextByte = stream.ReadByte();
                        if (nextByte != -1)
                        {
                            byte[] temp = new byte[readBuffer.Length * 2];
                            Buffer.BlockCopy(readBuffer, 0, temp, 0, readBuffer.Length);
                            Buffer.SetByte(temp, totalBytesRead, (byte)nextByte);
                            readBuffer = temp;
                            totalBytesRead++;
                        }
                    }
                }

                byte[] buffer = readBuffer;
                if (readBuffer.Length != totalBytesRead)
                {
                    buffer = new byte[totalBytesRead];
                    Buffer.BlockCopy(readBuffer, 0, buffer, 0, totalBytesRead);
                }
                return buffer;
            }
            finally
            {
                if (stream.CanSeek)
                {
                    stream.Position = originalPosition;
                }
            }
        }
    }
}
