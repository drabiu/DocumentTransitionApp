using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

using Service = DocumentTransitionUniversalApp.TransitionAppServices;
using Windows.UI.Popups;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace DocumentTransitionUniversalApp
{
	/// <summary>
	/// An empty page that can be used on its own or navigated to within a Frame.
	/// </summary>
	public sealed partial class MainPage : Page
	{
		StorageFile DocxFile;
		StorageFile XmlFile;

		public MainPage()
		{
			this.InitializeComponent();
		}

		private async void buttonDocx_Click(object sender, RoutedEventArgs e)
		{
			var picker = new FileOpenPicker();
			picker.ViewMode = PickerViewMode.List;
			picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
			picker.FileTypeFilter.Add(".docx");

			StorageFile file = await picker.PickSingleFileAsync();
			if (file != null)
			{
				DocxFile = file;
			}

			EnableSplitButton();
		}

		private async void buttonXml_Click(object sender, RoutedEventArgs e)
		{
			var picker = new FileOpenPicker();
			picker.ViewMode = PickerViewMode.List;
			picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
			picker.FileTypeFilter.Add(".xml");

			StorageFile file = await picker.PickSingleFileAsync();
			if (file != null)
			{
				XmlFile = file;
			}

			EnableSplitButton();
		}

		private async void buttonSplit_Click(object sender, RoutedEventArgs e)
		{
			Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			byte[] docxBinary = await StorageFileToByteArray(DocxFile);
			byte[] xmlBinary = await StorageFileToByteArray(XmlFile);
			var result = await serviceClient.SplitDocumentAsync(Path.GetFileNameWithoutExtension(DocxFile.Name), docxBinary, xmlBinary);
			SaveFiles(result);
			EnableMergeButton();
		}

		private async void buttonMerge_Click(object sender, RoutedEventArgs e)
		{
			Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			var files = await GetFiles();
            var result = await serviceClient.MergeDocumentAsync(files);
			SaveFile(result);
        }

		private async void SaveFiles(Service.SplitDocumentResponse response)
		{
			FolderPicker folderPicker = new FolderPicker();
			folderPicker.SuggestedStartLocation = PickerLocationId.Downloads;
			StorageFolder folder = await folderPicker.PickSingleFolderAsync();
			StorageFolder filesSaveFolder;
			try
			{
				filesSaveFolder = await folder.GetFolderAsync("Split Files");
			}
			catch (FileNotFoundException ex)
			{
				filesSaveFolder = await folder.CreateFolderAsync("Split Files");
			}

			foreach (Service.PersonFiles file in response.Body.SplitDocumentResult)
			{
				if (file.Person == "/")
				{
					StorageFile newFile;
                    try
					{
						newFile = await folder.GetFileAsync(file.Name);
					}
					catch (FileNotFoundException ex)
					{
						newFile = await filesSaveFolder.CreateFileAsync(file.Name);
					}
						
					using (var s = await newFile.OpenStreamForWriteAsync())
					{
						s.Write(file.Data, 0, file.Data.Length);
					}
				}
				else
				{
					StorageFolder currentSaveFolder;
					try
					{
						currentSaveFolder = await filesSaveFolder.GetFolderAsync(file.Person);
					}
					catch (FileNotFoundException ex)
					{
						currentSaveFolder = await filesSaveFolder.CreateFolderAsync(file.Person);
					}

					StorageFile newFile;
                    try
					{
						newFile = await currentSaveFolder.GetFileAsync(file.Name + ".docx");
					}
					catch (FileNotFoundException ex)
					{
						newFile = await currentSaveFolder.CreateFileAsync(file.Name + ".docx");
					}

					using (var s = await newFile.OpenStreamForWriteAsync())
					{
						s.Write(file.Data, 0, file.Data.Length);
					}
				}
			}
		}

		private async void SaveFile(Service.MergeDocumentResponse response)
		{			
		}

		private async Task<ObservableCollection<Service.PersonFiles>> GetFiles()
		{
			ObservableCollection<Service.PersonFiles> files = new ObservableCollection<Service.PersonFiles>();
			FolderPicker folderPicker = new FolderPicker();
			folderPicker.SuggestedStartLocation = PickerLocationId.Downloads;
			StorageFolder folder = await folderPicker.PickSingleFolderAsync();
			StorageFile xmlFile;
			try
			{
				xmlFile = await folder.GetFileAsync("mergeXmlDefinition.xml");
			}
			catch (FileNotFoundException ex)
			{
				var dialog = new MessageDialog("mergeXmlDefinition.xml does not exist");
				await dialog.ShowAsync();
				return files;
			}

			StorageFile templateFile;
			try
			{
				templateFile = await folder.GetFileAsync("template.docx");
			}
			catch (FileNotFoundException ex)
			{
				var dialog = new MessageDialog("template.docx does not exist");
				await dialog.ShowAsync();
				return files;
			}

			foreach (StorageFolder subFolder in await folder.GetFoldersAsync())
			{
				foreach (StorageFile fileToLoad in await subFolder.GetFilesAsync())
				{
					var personFile = new Service.PersonFiles();
					personFile.Name = Path.GetFileNameWithoutExtension(fileToLoad.Name);
					personFile.Data = await StorageFileToByteArray(fileToLoad);
					personFile.Person = subFolder.Name;
					files.Add(personFile);
                }
			}

			var personXmlFile = new Service.PersonFiles();
			personXmlFile.Name = xmlFile.Name;
			personXmlFile.Data = await StorageFileToByteArray(xmlFile);
			personXmlFile.Person = "/";

			files.Add(personXmlFile);

			var personTemplateFile = new Service.PersonFiles();
			personTemplateFile.Name = templateFile.Name;
			personTemplateFile.Data = await StorageFileToByteArray(templateFile);
			personTemplateFile.Person = "/";

			files.Add(personTemplateFile);

			return files;
		}

		private void EnableSplitButton()
		{
			if (DocxFile != null && XmlFile != null)
				buttonSplit.IsEnabled = true;
		}

		private void EnableMergeButton()
		{
			buttonMerge.IsEnabled = true;
		}

		private async Task<byte[]> StorageFileToByteArray(StorageFile file)
		{
			var fileStream = await file.OpenAsync(FileAccessMode.Read);
			return ReadFully(fileStream.AsStream());
		}

		public static byte[] ReadFully(Stream input)
		{
			byte[] buffer = new byte[16 * 1024];
			using (MemoryStream ms = new MemoryStream())
			{
				int read;
				while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
				{
					ms.Write(buffer, 0, read);
				}
				return ms.ToArray();
			}
		}
	}
}
