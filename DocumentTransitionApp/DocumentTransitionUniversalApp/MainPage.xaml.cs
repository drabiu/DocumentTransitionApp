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

using DocumentTransitionUniversalApp.Data_Structures;
using Service = DocumentTransitionUniversalApp.TransitionAppServices;
using DocumentTransitionUniversalApp.Views;
using Windows.UI.Popups;
using Windows.UI.Core;
using System.Collections.Generic;
using System.Globalization;
using DocumentTransitionUniversalApp.Helpers;

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
		public string FileName;
		public DocumentType FileType;
        public ElementTypes DocumentElementTypes;

        public Frame AppFrame { get { return this.frame; } }
		public byte[] documentBinary;
		public byte[] xmlBinary;

        public enum DocumentType
        {
            Word,
            Excel,
            Presentation
        }

        private bool _wasSplit;
        private bool _wasEditParts;
        private StorageFile XmlFile;

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
            ResetControls();

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
                        DocumentElementTypes = new WordElementType();
                        break;
					case (".xlsx"):
						FileType = DocumentType.Excel;
                        DocumentElementTypes = new ExcelElementType();
                        break;
					case (".pptx"):
						FileType = DocumentType.Presentation;
                        DocumentElementTypes = new PresentationElementType();
						break;
				}

				documentBinary = await StorageFileToByteArray(DocumentFile);
			}

            InitButtons();
        }

		private async void buttonXml_Click(object sender, RoutedEventArgs e)
		{
            _wasEditParts = true;
			this.AppFrame.Navigate(typeof(WordSelectPartsPage), this);
        }

		private async void buttonSplit_Click(object sender, RoutedEventArgs e)
		{
            Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
            string extension = string.Empty;
            var result = new ObservableCollection<Service.PersonFiles>();
            try
            {
                switch (FileType)
                {
                    case (DocumentType.Word):
                        extension = ".docx";
                        var response = await serviceClient.SplitDocumentAsync(Path.GetFileNameWithoutExtension(FileName), documentBinary, xmlBinary);
                        result = response.Body.SplitDocumentResult;
                        break;
                    case (DocumentType.Excel):
                        extension = ".xlsx";
                        break;
                    case (DocumentType.Presentation):
                        extension = ".pptx";
                        var response2 = await serviceClient.SplitPresentationAsync(Path.GetFileNameWithoutExtension(FileName), documentBinary, xmlBinary);
                        result = response2.Body.SplitPresentationResult;
                        break;
                }

                SaveFiles(result, extension);
                _wasSplit = true;
            }
            catch (System.ServiceModel.CommunicationException ex)
            {
                MessageDialog dialog = new MessageDialog(ex.Message);
                await dialog.ShowAsync();
            }
        }

		private async void buttonMerge_Click(object sender, RoutedEventArgs e)
		{
			Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			var files = await GetFiles();
            if (files.Any())
            {
                try
                {
                    var result = await serviceClient.MergeDocumentAsync(files);
                    FileHelper.SaveFile(result.Body.MergeDocumentResult, "Merged Document Name", ".docx", ".xlsx", ".pptx");
                }
                catch (System.ServiceModel.CommunicationException ex)
                {
                    MessageDialog dialog = new MessageDialog(ex.Message);
                    await dialog.ShowAsync();
                }
            }
        }

        private async void SaveFiles(ObservableCollection<Service.PersonFiles> personFiles, string fileExtension)
        {
            FolderPicker folderPicker = new FolderPicker();
            folderPicker.SuggestedStartLocation = PickerLocationId.Downloads;
            folderPicker.FileTypeFilter.Add(".docx");
            folderPicker.FileTypeFilter.Add(".xlsx");
            folderPicker.FileTypeFilter.Add(".pptx");
            StorageFolder folder = await folderPicker.PickSingleFolderAsync();
            StorageFolder filesSaveFolder;
            if (folder != null)
            { 
                try
                {
                    filesSaveFolder = await folder.GetFolderAsync(string.Format("Split Files ({0})", FileName));
                    await filesSaveFolder.DeleteAsync();
                }
                catch (FileNotFoundException ex)
                {
                }
                finally
                {
                    filesSaveFolder = await folder.CreateFolderAsync(string.Format("Split Files ({0})", FileName));
                }

                foreach (Service.PersonFiles file in personFiles)
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
                            newFile = await currentSaveFolder.GetFileAsync(file.Name + fileExtension);
                        }
                        catch (FileNotFoundException ex)
                        {
                            newFile = await currentSaveFolder.CreateFileAsync(file.Name + fileExtension);
                        }

                        using (var s = await newFile.OpenStreamForWriteAsync())
                        {
                            s.Write(file.Data, 0, file.Data.Length);
                        }
                    }
                }
            }
		}
   
		private async Task<ObservableCollection<Service.PersonFiles>> GetFiles()
		{
			ObservableCollection<Service.PersonFiles> files = new ObservableCollection<Service.PersonFiles>();
			FolderPicker folderPicker = new FolderPicker();
            folderPicker.FileTypeFilter.Add(".docx");
            folderPicker.SuggestedStartLocation = PickerLocationId.Downloads;
			StorageFolder folder = await folderPicker.PickSingleFolderAsync();
            if (folder != null)
            {
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
            }

			return files;
		}

        private void ResetControls()
        {
            WordPartPage = null;
            _wasSplit = false;
            _wasEditParts = false;
            DocumentFile = null;
            XmlFile = null;
        }

		private void EnablePartsButton()
		{
			buttonXml.IsEnabled = DocumentFile != null;
        }

		private void EnableSplitButton()
		{
			buttonSplit.IsEnabled = DocumentFile != null && XmlFile != null;
		}

        private void EnableGenerateButton()
        {
            buttonGenerateSplit.IsEnabled = _wasEditParts;
        }

        private void EnableLoadButton()
        {
            buttonLoadSplit.IsEnabled = DocumentFile != null;
        }
		
		private void InitButtons()
		{
			EnablePartsButton();
			EnableSplitButton();
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
                this.xmlBinary = main.xmlBinary;

				InitButtons();
			}
		}

        private async void buttonGenerateSplit_Click(object sender, RoutedEventArgs e)
        {
            Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
            var selectionParts = new ObservableCollection<Service.PartsSelectionTreeElement>();
            foreach (var part in WordPartPage._pageData.SelectionParts)
                selectionParts.Add(part.ConvertToPartsSelectionTreeElement());

            try
            {
                var result = await serviceClient.GenerateSplitDocumentAsync(Path.GetFileNameWithoutExtension(FileName), selectionParts);
                string splitFileName = string.Format("split_{1}_{0}.xml", DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture), FileName);
                FileHelper.SaveFile(result.Body.GenerateSplitDocumentResult, splitFileName, ".xml");
            }
            catch (System.ServiceModel.CommunicationException ex)
            {
                MessageDialog dialog = new MessageDialog(ex.Message);
                await dialog.ShowAsync();
            }
        }

        private async void buttonLoadSplit_Click(object sender, RoutedEventArgs e)
        {
            try
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
                    List<PartsSelectionTreeElement<ElementTypes>> parts = new List<PartsSelectionTreeElement<ElementTypes>>();
                    var response = await GetPartsFromXml();
                    if (!response.IsError)
                    {
                        var partsFromXml = response.Data as ObservableCollection<Service.PartsSelectionTreeElement>;
                        foreach (var element in partsFromXml)
                        {
                            var item = new PartsSelectionTreeElement<ElementTypes>(element.Id, element.ElementId, DocumentElementTypes, element.Name, element.Indent, element.Selected, element.OwnerName);
                            parts.Add(item);
                        }

                        var names = partsFromXml.Select(p => p.OwnerName).Where(n => !string.IsNullOrEmpty(n)).Distinct().ToList();
                        List<Data_Structures.ComboBoxItem> comboItems = new List<Data_Structures.ComboBoxItem>();
                        int indexer = 1;
                        foreach (var name in names)
                            comboItems.Add(new Data_Structures.ComboBoxItem() { Id = indexer++, Name = name });

                        WordPartPage = new WordSelectPartsPage();
                        Data_Structures.WordPartsPageData pageData = new Data_Structures.WordPartsPageData();
                        pageData.SelectionParts = parts;
                        pageData.LastId = names.Count();
                        pageData.ComboItems.AddRange(comboItems);
                        WordPartPage.CopyDataToControl(pageData);
                        EnableSplitButton();
                    }
                    else
                    {
                        MessageDialog dialog = new MessageDialog(response.Message);
                        await dialog.ShowAsync();
                    }                  
                }
            }
            catch(Exception ex)
            {
                MessageDialog dialog = new MessageDialog(ex.Message);
                await dialog.ShowAsync();
            }           
        }

        private async Task<Service.ServiceResponse> GetPartsFromXml()
        {
            var result = new TransitionAppServices.ServiceResponse();
            Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();

            try
            {
                switch (FileType)
                {
                    case (MainPage.DocumentType.Word):
                        var response = await serviceClient.GetDocumentPartsFromXmlAsync(FileName, documentBinary, xmlBinary);
                        result = response.Body.GetDocumentPartsFromXmlResult;
                        break;
                    case (MainPage.DocumentType.Excel):
                        break;
                    case (MainPage.DocumentType.Presentation):
                        var response2 = await serviceClient.GetPresentationPartsFromXmlAsync(FileName, documentBinary, xmlBinary);
                        result = response2.Body.GetPresentationPartsFromXmlResult;
                        break;
                }
            }
            catch (System.ServiceModel.CommunicationException ex)
            {
                MessageDialog dialog = new MessageDialog(ex.Message);
                await dialog.ShowAsync();
            }

            return result;
        }
    }
}
