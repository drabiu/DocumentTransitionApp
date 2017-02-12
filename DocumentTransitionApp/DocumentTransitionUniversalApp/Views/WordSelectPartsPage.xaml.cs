using System;
using System.Collections.Generic;
using System.Linq;
using Windows.UI;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Windows.UI.Popups;
using System.Collections.ObjectModel;
using Service = DocumentTransitionUniversalApp.TransitionAppServices;
using DocumentTransitionUniversalApp.Data_Structures;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace DocumentTransitionUniversalApp.Views
{
	/// <summary>
	/// An empty page that can be used on its own or navigated to within a Frame.
	/// </summary>
	public sealed partial class WordSelectPartsPage : Page
	{
		MainPage _source;
        public WordPartsPageData _pageData { get; private set; }

		public WordSelectPartsPage()
		{
			this.InitializeComponent();
            InitializeVariables();
        }

        public void CopyDataToControl(WordPartsPageData data)
        {
            _pageData = data;
        }

		protected override void OnNavigatedTo(NavigationEventArgs e)
		{
			if (e.Parameter is MainPage)
			{
				this._source = e.Parameter as MainPage;
                if (this._source.WordPartPage != null)
                {
                    _pageData = new WordPartsPageData(this._source.WordPartPage);
                    InitializeVariables();
                }
                else
                {
                    InitializeItems();
                }
			}

            base.OnNavigatedTo(e);
		}

		private void BackButton_Click(object sender, RoutedEventArgs e)
		{
            _source.WordPartPage = this;
            this.Frame.Navigate(typeof(MainPage), _source);
		}

        private void PrepareListOfItems(ObservableCollection<Service.PartsSelectionTreeElement> elements, ElementTypes elementType)
        {
            foreach (var element in elements)
            {
                var item = new PartsSelectionTreeElement<ElementTypes>(element.Id, element.ElementId, elementType, element.Name, element.Indent);
                _pageData.SelectionParts.Add(item);
            }
        }

		private void CreateSelectPartsUI(IEnumerable<PartsSelectionTreeElement<ElementTypes>> elements)
		{
            WordSelectPartsItems.Items.Clear();
            foreach (var element in elements)
            {
                CreateButtonBlock(element);
            }
        }

        private void CreateButtonBlock(PartsSelectionTreeElement<ElementTypes> element)
		{
            Button button = new Button();

            if (element.Selected)
                button.Background = new SolidColorBrush(Colors.Honeydew);
            else
                button.Background = new SolidColorBrush(Colors.Transparent);

            button.Name = element.Id;
            button.Margin = new Thickness(element.Indent * 20, 5, 0, 5);
            button.Content = element.Name;

            ToolTip toolTip = new ToolTip();
            toolTip.Content = element.Name;
            ToolTipService.SetToolTip(button, toolTip);

            if ((string)comboBox.SelectedItem != null)
            {
                var comboItem = Data_Structures.ComboBoxItem.GetComboBoxItemByName(_pageData.ComboItems, (string)comboBox.SelectedItem);
                if (element.CheckIfCanBeSelected(Data_Structures.ComboBoxItem.GetComboBoxItemByName(_pageData.ComboItems, (string)comboItem.Name).Name) && comboItem.Id != WordPartsPageData.AllItemsId)
                {
                    button.Tapped += Button_Tapped;
                }
                else
                    button.Background = new SolidColorBrush(Colors.DimGray);
            }                               
            else
                button.Background = new SolidColorBrush(Colors.DimGray);

            WordSelectPartsItems.Items.Add(button);
        }

        private void Button_Tapped(object sender, TappedRoutedEventArgs e)
        {
            var ownerName = Data_Structures.ComboBoxItem.GetComboBoxItemByName(_pageData.ComboItems, (string)comboBox.SelectedItem).Name;
            var button = sender as Button;
            var selectedElement = _pageData.SelectionParts.Single(el => el.Id == button.Name);
            if (selectedElement.Selected)
            {
                selectedElement.SelectItem(string.Empty);
                button.Background = new SolidColorBrush(Colors.White);
            }
            else
            {
                selectedElement.SelectItem(ownerName);
                button.Background = new SolidColorBrush(Colors.Honeydew);
            }
        }

        private async void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
            CreateSelectPartsUI(_pageData.SelectionParts);
		}

        private void PersonTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox.Text.Length > 0)
                AddButton.IsEnabled = true;
            else
                AddButton.IsEnabled = false;
        }

        private async void AddButton_Click(object sender, RoutedEventArgs e)
        {
            var name = PersonTextBox.Text;
            if (_pageData.ComboItems.Any(it => it.Name == name))
            {
                var dialog = new MessageDialog("There is already a person with this name");
                await dialog.ShowAsync();
            }
            else
            {
                _pageData.ComboItems.Add(new Data_Structures.ComboBoxItem() { Id = ++_pageData.LastId, Name = name });
            }
            
            comboBox.ItemsSource = _pageData.ComboItems.Select(cb => cb.Name);
            comboBox.SelectedValue = name;
            PersonTextBox.Text = string.Empty;
        }

        private void InitializeVariables()
        {
            if (_pageData == null)
                _pageData = new WordPartsPageData();

            comboBox.ItemsSource = _pageData.ComboItems.Select(cb => cb.Name);
            AddButton.IsEnabled = false;
        }

        private async void InitializeItems()
        {
            switch (_source.FileType)
            {
                case (MainPage.DocumentType.Word):
                    InitWord();
                    break;
                case (MainPage.DocumentType.Excel):
                    InitExcel();
                    break;
                case (MainPage.DocumentType.Presentation):
                    InitPresentation();
                    break;
            }     
        }

        private async void InitWord()
        {
            Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
            try
            {
                var result = await serviceClient.GetDocumentPartsAsync(_source.FileName, _source.documentBinary);
                PrepareListOfItems(result.Body.GetDocumentPartsResult, _source.DocumentElementTypes);
            }
            catch(System.ServiceModel.CommunicationException ex)
            {
                MessageDialog dialog = new MessageDialog(ex.Message);
                await dialog.ShowAsync();
            }           
        }

        private async void InitPresentation()
        {
            Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
            var result = await serviceClient.GetPresentationPartsAsync(_source.FileName, _source.documentBinary);
            PrepareListOfItems(result.Body.GetPresentationPartsResult, _source.DocumentElementTypes);
        }

        private async void InitExcel()
        {
            throw new NotImplementedException();
        }
    }    
}
