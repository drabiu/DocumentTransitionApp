using System;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;

using Service = DocumentTransitionUniversalApp.TransitionAppServices;
using DocumentTransitionUniversalApp.Views;
using Windows.UI.Popups;
using Windows.UI.Core;
using System.Collections.Generic;
using System.Globalization;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace DocumentTransitionUniversalApp
{
	/// <summary>
	/// An empty page that can be used on its own or navigated to within a Frame.
	/// </summary>
	public sealed partial class MainPage : Page
	{
        public WordSelectPartsPage WordPartPage;
		public StorageFile DocumentFile;
		StorageFile XmlFile;
		public string FileName;
		DocumentType FileType;
		bool _wasSplit;
        bool _wasEditParts;
		public Frame AppFrame { get { return this.frame; } }
		public byte[] documentBinary;
		public byte[] xmlBinary;

		public enum DocumentType
		{
			Word,
			Excel,
			Presentation
		}

		public MainPage()
		{
			this.InitializeComponent();

			SystemNavigationManager.GetForCurrentView().BackRequested += SystemNavigationManager_BackRequested;
		}

		private void SystemNavigationManager_BackRequested(object sender, BackRequestedEventArgs e)
		{
			bool handled = e.Handled;
			this.BackRequested(ref handled);
			e.Handled = handled;
		}

		private void BackRequested(ref bool handled)
		{
			// Get a hold of the current frame so that we can inspect the app back stack.

			if (this.AppFrame == null)
				return;

			// Check to see if this is the top-most page on the app back stack.
			if (this.AppFrame.CanGoBack && !handled)
			{
				// If not, set the event to handled and go back to the previous page in the app.
				handled = true;
				this.AppFrame.GoBack();
			}
		}

		private async void buttonDocx_Click(object sender, RoutedEventArgs e)
		{
			var picker = new FileOpenPicker();
			picker.ViewMode = PickerViewMode.List;
			picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
			picker.FileTypeFilter.Add(".docx");
			picker.FileTypeFilter.Add(".xlsx");
			picker.FileTypeFilter.Add(".pptx");

			StorageFile file = await picker.PickSingleFileAsync();
			if (file != null)
			{
				DocumentFile = file;
				FileName = DocumentFile.Name;
				switch (Path.GetExtension(file.Name))
				{
					case (".docx"):
						FileType = DocumentType.Word;
						break;
					case (".xlsx"):
						FileType = DocumentType.Excel;
						break;
					case (".pptx"):
						FileType = DocumentType.Presentation;
						break;
				}

				documentBinary = await StorageFileToByteArray(DocumentFile);
			}

			EnablePartsButton();
            EnableLoadButton();
        }

		private async void buttonXml_Click(object sender, RoutedEventArgs e)
		{
            _wasEditParts = true;
			switch (FileType)
			{
				case (DocumentType.Word):
					this.AppFrame.Navigate(typeof(WordSelectPartsPage), this);
					break;
				case (DocumentType.Excel):
					this.AppFrame.Navigate(typeof(ExcelSelectPartsPage), this);
					break;
				case (DocumentType.Presentation):
					this.AppFrame.Navigate(typeof(PresentationSelectPartsPage), this);
					break;
			}
        }

		private async void buttonSplit_Click(object sender, RoutedEventArgs e)
		{
			Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			var result = await serviceClient.SplitDocumentAsync(Path.GetFileNameWithoutExtension(FileName), documentBinary, xmlBinary);
			SaveFiles(result);
			_wasSplit = true;
			EnableMergeButton();
		}

		private async void buttonMerge_Click(object sender, RoutedEventArgs e)
		{
			Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			var files = await GetFiles();
            var result = await serviceClient.MergeDocumentAsync(FileName, files);
			SaveFile(result.Body.MergeDocumentResult, FileName, ".docx", ".xlsx", ".pptx");
        }

		private async void SaveFiles(Service.SplitDocumentResponse response)
		{
			FolderPicker folderPicker = new FolderPicker();
			folderPicker.SuggestedStartLocation = PickerLocationId.Downloads;
            folderPicker.FileTypeFilter.Add(".docx");
            folderPicker.FileTypeFilter.Add(".xlsx");
            folderPicker.FileTypeFilter.Add(".pptx");
            StorageFolder folder = await folderPicker.PickSingleFolderAsync();
			StorageFolder filesSaveFolder;
			try
			{
				filesSaveFolder = await folder.GetFolderAsync("Split Files");
				await filesSaveFolder.DeleteAsync();
			}
			catch (FileNotFoundException ex)
			{			
			}
			finally
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

        private async void SaveFile(byte[] fileBinary, string fileName, params string[] filters)
        {
            FolderPicker folderPicker = new FolderPicker();
            foreach (var filter in filters)
                folderPicker.FileTypeFilter.Add(filter);

            folderPicker.SuggestedStartLocation = PickerLocationId.Downloads;
            StorageFolder folder = await folderPicker.PickSingleFolderAsync();
            StorageFile newFile;
            try
            {
                newFile = await folder.GetFileAsync(fileName);
            }
            catch (FileNotFoundException ex)
            {
                newFile = await folder.CreateFileAsync(fileName);
            }

            using (var s = await newFile.OpenStreamForWriteAsync())
            {
                s.Write(fileBinary, 0, fileBinary.Length);
            }
        }

		private async Task<ObservableCollection<Service.PersonFiles>> GetFiles()
		{
			ObservableCollection<Service.PersonFiles> files = new ObservableCollection<Service.PersonFiles>();
			FolderPicker folderPicker = new FolderPicker();
            folderPicker.FileTypeFilter.Add(".docx");
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

		private void EnablePartsButton()
		{
			if (DocumentFile != null)
				buttonXml.IsEnabled = true;
        }

		private void EnableSplitButton()
		{
			if (DocumentFile != null && XmlFile != null)
				buttonSplit.IsEnabled = true;
		}

		private void EnableMergeButton()
		{
			if (_wasSplit)
				buttonMerge.IsEnabled = true;
		}

        private void EnableGenerateButton()
        {
            if (_wasEditParts)
                buttonGenerateSplit.IsEnabled = true;
        }

        private void EnableLoadButton()
        {
            if (DocumentFile != null)
                buttonLoadSplit.IsEnabled = true;
        }
		
		private void InitButtons()
		{
			EnablePartsButton();
			EnableSplitButton();
			EnableMergeButton();
            EnableGenerateButton();
            EnableLoadButton();
        }

		public async Task<byte[]> StorageFileToByteArray(StorageFile file)
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

		protected override void OnNavigatedTo(NavigationEventArgs e)
		{
			if (e.Parameter is MainPage)
			{
				var main = e.Parameter as MainPage;
				this.DocumentFile = main.DocumentFile;
                this.documentBinary = main.documentBinary;
				this.XmlFile = main.XmlFile;
				this.FileName = main.FileName;
				this.FileType = main.FileType;
				this._wasSplit = main._wasSplit;
                this._wasEditParts = main._wasEditParts;
                this.WordPartPage = main.WordPartPage;

				InitButtons();
			}
		}

        private async void buttonGenerateSplit_Click(object sender, RoutedEventArgs e)
        {
            Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
            var selectionParts = new ObservableCollection<DocumentTransitionUniversalApp.TransitionAppServices.PartsSelectionTreeElement>();
            foreach (var part in WordPartPage._pageData.SelectionParts)
                selectionParts.Add(part.ConvertToPartsSelectionTreeElement());
            
            var result = await serviceClient.GenerateSplitDocumentAsync(Path.GetFileNameWithoutExtension(FileName), selectionParts);
            string splitFileName = string.Format("split{0}.xml", DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture));
            SaveFile(result.Body.GenerateSplitDocumentResult, splitFileName, ".xml");
        }

        private async void buttonLoadSplit_Click(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            picker.ViewMode = PickerViewMode.List;
            picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            picker.FileTypeFilter.Add(".xml");

            StorageFile file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                XmlFile = file;
                xmlBinary = await StorageFileToByteArray(XmlFile);
                EnableSplitButton();
                Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
                var result = await serviceClient.GetDocumentPartsFromXmlAsync(Path.GetFileNameWithoutExtension(FileName), documentBinary, xmlBinary);
                List<PartsSelectionTreeElement<ElementTypes.WordElementType>> parts = new List<PartsSelectionTreeElement<ElementTypes.WordElementType>>();
                foreach (var element in result.Body.GetDocumentPartsFromXmlResult)
                {
                    var item = new PartsSelectionTreeElement<ElementTypes.WordElementType>(element.Id, element.ElementId, ElementTypes.WordElementType.Paragraph, element.Name, element.Indent, element.Selected);
                    parts.Add(item);
                }

                
                var names = result.Body.GetDocumentPartsFromXmlResult.Select(p => p.OwnerName).Where(n => !string.IsNullOrEmpty(n)).Distinct().ToList();
                List<Views.ComboBoxItem> comboItems = new List<Views.ComboBoxItem>();
                int indexer = 1;
                foreach (var name in names)
                    comboItems.Add(new Views.ComboBoxItem() { Id = indexer++, Name = name });

                WordPartPage = new WordSelectPartsPage();
                WordPartsPageData pageData = new WordPartsPageData();
                pageData.SelectionParts = parts;
                pageData.LastId = names.Count();
                pageData.ComboItems.AddRange(comboItems);
                WordPartPage.CopyDataToControl(pageData);
            }           
        }
    }
}
