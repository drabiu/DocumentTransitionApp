using DocumentTransitionUniversalApp.Data_Structures;
using DocumentTransitionUniversalApp.Helpers;
using DocumentTransitionUniversalApp.Repositories;
using DocumentTransitionUniversalApp.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Core;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;
using Service = DocumentTransitionUniversalApp.TransitionAppServices;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace DocumentTransitionUniversalApp
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        #region Fields

        public WordSelectPartsPage WordPartPage;
        public StorageFile DocumentFile;
        private string _fileName;
        public string FileName
        {
            get
            {
                return _fileName;
            }
            set
            {
                loadedDocument.Text = value;
                _fileName = value;
            }
        }
        public DocumentType FileType;

        public Frame AppFrame { get { return this.frame; } }
        public byte[] documentBinary;
        public byte[] xmlBinary;
        public static ServiceDecorator Service = new ServiceDecorator();
        ServiceRepository _serviceRepo;

        public enum DocumentType
        {
            Word,
            Excel,
            Presentation
        }

        private bool _wasSplit;
        private bool _wasEditParts;
        private StorageFile XmlFile;

        #endregion

        public MainPage()
        {
            this.InitializeComponent();

            SystemNavigationManager.GetForCurrentView().BackRequested += SystemNavigationManager_BackRequested;

            _serviceRepo = new ServiceRepository();
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

        #region Events

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
                SetFileType(file.Name);
                documentBinary = await StorageFileToByteArray(DocumentFile);
            }

            InitButtons();
        }

        private void buttonXml_Click(object sender, RoutedEventArgs e)
        {
            _wasEditParts = true;
            this.AppFrame.Navigate(typeof(WordSelectPartsPage), this);
        }

        private async void buttonSplit_Click(object sender, RoutedEventArgs e)
        {
            string extension = string.Empty;
            ObservableCollection<TransitionAppWCFSerivce.PersonFiles> result = new ObservableCollection<TransitionAppWCFSerivce.PersonFiles>();
            try
            {
                switch (FileType)
                {
                    case (DocumentType.Word):
                        extension = ".docx";
                        result = await _serviceRepo.SplitWordAsync(Path.GetFileNameWithoutExtension(FileName), documentBinary, xmlBinary);
                        break;
                    case (DocumentType.Excel):
                        extension = ".xlsx";
                        result = await _serviceRepo.SplitExcelAsync(Path.GetFileNameWithoutExtension(FileName), documentBinary, xmlBinary);
                        break;
                    case (DocumentType.Presentation):
                        extension = ".pptx";
                        result = await _serviceRepo.SplitPresentationAsync(Path.GetFileNameWithoutExtension(FileName), documentBinary, xmlBinary);
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
            var files = await GetFiles();
            if (files.Any())
            {
                try
                {
                    byte[] result;
                    switch (FileType)
                    {
                        case (DocumentType.Word):
                            result = await _serviceRepo.MergeWordAsync(files);
                            FileHelper.SaveFile(result, "Merged Document Name", ".docx");
                            break;
                        case (DocumentType.Excel):
                            result = await _serviceRepo.MergeExcelAsync(files);
                            FileHelper.SaveFile(result, "Merged Document Name", ".xlsx");
                            break;
                        case (DocumentType.Presentation):
                            result = await _serviceRepo.MergePresentationAsync(files);
                            FileHelper.SaveFile(result, "Merged Document Name", ".pptx");
                            break;
                    }
                }
                catch (System.ServiceModel.CommunicationException ex)
                {
                    MessageDialog dialog = new MessageDialog(ex.Message);
                    await dialog.ShowAsync();
                }
            }
        }

        private async void buttonSettings_Click(object sender, RoutedEventArgs e)
        {
            var endpointAdress = await InputTextDialogAsync("Set service adress", Service.DefaultEndpoint);
            if (!string.IsNullOrEmpty(endpointAdress))
                Service.DefaultEndpoint = endpointAdress;
        }

        private async void buttonGenerateSplit_Click(object sender, RoutedEventArgs e)
        {
            var selectionParts = new ObservableCollection<TransitionAppWCFSerivce.PartsSelectionTreeElement>();
            foreach (var part in WordPartPage._pageData.SelectionParts)
                selectionParts.Add(part.ConvertToServicePartsSelectionTreeElement());

            string splitFileName = string.Format("split_{1}_{0}.xml", DateTime.UtcNow.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture), FileName);
            try
            {
                byte[] result;
                switch (FileType)
                {
                    case (DocumentType.Word):
                        result = await _serviceRepo.GenerateSplitWordAsync(Path.GetFileNameWithoutExtension(FileName), selectionParts);
                        FileHelper.SaveFile(result, splitFileName, ".xml");
                        break;
                    case (DocumentType.Excel):
                        result = await _serviceRepo.GenerateSplitExcelAsync(Path.GetFileNameWithoutExtension(FileName), selectionParts);
                        FileHelper.SaveFile(result, splitFileName, ".xml");
                        break;
                    case (DocumentType.Presentation):
                        result = await _serviceRepo.GenerateSplitPresentationAsync(Path.GetFileNameWithoutExtension(FileName), selectionParts);
                        FileHelper.SaveFile(result, splitFileName, ".xml");
                        break;
                }

                EnableSplitButton();
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
                    ObservableCollection<PartsSelectionTreeElement> parts = new ObservableCollection<PartsSelectionTreeElement>();
                    var response = await GetPartsFromXml();
                    if (!response.IsError)
                    {
                        var partsFromXml = response.Data as ObservableCollection<TransitionAppWCFSerivce.PartsSelectionTreeElement>;
                        foreach (var element in partsFromXml)
                        {
                            var item = PartsSelectionTreeElement.ConvertToPartsSelectionTreeElement(element);
                            parts.Add(item);
                        }

                        var names = partsFromXml.Select(p => p.OwnerName).Where(n => !string.IsNullOrEmpty(n)).Distinct().ToList();
                        List<Data_Structures.ComboBoxItem> comboItems = new List<Data_Structures.ComboBoxItem>();
                        int indexer = 1;
                        foreach (var name in names)
                            comboItems.Add(new Data_Structures.ComboBoxItem() { Id = indexer++, Name = name });

                        WordPartPage = new WordSelectPartsPage();
                        WordPartsPageData pageData = new WordPartsPageData();
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
            catch (Exception ex)
            {
                MessageDialog dialog = new MessageDialog(ex.Message);
                await dialog.ShowAsync();
            }
        }

        #endregion

        #region Private methods

        private async void SaveFiles(ObservableCollection<TransitionAppWCFSerivce.PersonFiles> personFiles, string fileExtension)
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

                foreach (TransitionAppWCFSerivce.PersonFiles file in personFiles)
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

        private async Task<ObservableCollection<TransitionAppWCFSerivce.PersonFiles>> GetFiles()
        {
            ObservableCollection<TransitionAppWCFSerivce.PersonFiles> files = new ObservableCollection<TransitionAppWCFSerivce.PersonFiles>();
            FolderPicker folderPicker = new FolderPicker();
            folderPicker.FileTypeFilter.Add(".docx");
            folderPicker.FileTypeFilter.Add(".pptx");
            folderPicker.FileTypeFilter.Add(".xlsx");
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

                string templateName = string.Empty;
                StorageFile templateFile;
                try
                {
                    IReadOnlyList<StorageFile> filesInFolder = await folder.GetFilesAsync();
                    foreach (StorageFile item in filesInFolder)
                    {
                        if (Path.GetFileNameWithoutExtension(item.Name) == "template")
                        {
                            templateName = item.Name;
                            break;
                        }
                    }

                    templateFile = await folder.GetFileAsync(templateName);
                    SetFileType(templateName);
                }
                catch (FileNotFoundException ex)
                {
                    var dialog = new MessageDialog(string.Format("{0} does not exist", templateName));
                    await dialog.ShowAsync();
                    return files;
                }
                catch (ArgumentException ex)
                {
                    var dialog = new MessageDialog(string.Format("template file does not exist", templateName));
                    await dialog.ShowAsync();
                    return files;
                }

                foreach (StorageFolder subFolder in await folder.GetFoldersAsync())
                {
                    foreach (StorageFile fileToLoad in await subFolder.GetFilesAsync())
                    {
                        var personFile = new TransitionAppWCFSerivce.PersonFiles();
                        personFile.Name = Path.GetFileNameWithoutExtension(fileToLoad.Name);
                        personFile.Data = await StorageFileToByteArray(fileToLoad);
                        personFile.Person = subFolder.Name;
                        files.Add(personFile);
                    }
                }

                var personXmlFile = new TransitionAppWCFSerivce.PersonFiles();
                personXmlFile.Name = xmlFile.Name;
                personXmlFile.Data = await StorageFileToByteArray(xmlFile);
                personXmlFile.Person = "/";

                files.Add(personXmlFile);

                var personTemplateFile = new TransitionAppWCFSerivce.PersonFiles();
                personTemplateFile.Name = templateFile.Name;
                personTemplateFile.Data = await StorageFileToByteArray(templateFile);
                personTemplateFile.Person = "/";

                files.Add(personTemplateFile);
            }

            return files;
        }

        private void SetFileType(string fileName)
        {
            switch (Path.GetExtension(fileName))
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
        }

        private void ResetControls()
        {
            WordPartPage = null;
            _wasSplit = false;
            _wasEditParts = false;
            DocumentFile = null;
            FileName = string.Empty;
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
            return StreamHelper.ReadFully(fileStream.AsStream());
        }

        private async Task<string> InputTextDialogAsync(string title, string textBoxContent)
        {
            TextBox inputTextBox = new TextBox();
            inputTextBox.AcceptsReturn = false;
            inputTextBox.Height = 32;
            inputTextBox.Text = textBoxContent;
            ContentDialog dialog = new ContentDialog();
            dialog.Content = inputTextBox;
            dialog.Title = title;
            dialog.IsSecondaryButtonEnabled = true;
            dialog.PrimaryButtonText = "Ok";
            dialog.SecondaryButtonText = "Cancel";
            if (await dialog.ShowAsync() == ContentDialogResult.Primary)
                return inputTextBox.Text;
            else
                return string.Empty;
        }

        private async Task<TransitionAppWCFSerivce.GetPartsFromXmlServiceResponse> GetPartsFromXml()
        {
            var result = new TransitionAppWCFSerivce.GetPartsFromXmlServiceResponse();
            try
            {
                switch (FileType)
                {
                    case (DocumentType.Word):
                        result = await _serviceRepo.GetWordPartsFromXmlAsync(FileName, documentBinary, xmlBinary);
                        break;
                    case (DocumentType.Excel):
                        result = await _serviceRepo.GetExcelPartsFromXmlAsync(FileName, documentBinary, xmlBinary);
                        break;
                    case (DocumentType.Presentation):
                        result = await _serviceRepo.GetPresentationPartsFromXmlAsync(FileName, documentBinary, xmlBinary);
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

        #endregion
    }
}
