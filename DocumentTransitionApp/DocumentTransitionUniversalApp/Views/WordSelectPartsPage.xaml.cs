using DocumentTransitionUniversalApp.Data_Structures;
using DocumentTransitionUniversalApp.Helpers;
using DocumentTransitionUniversalApp.Extension_Methods;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using Windows.UI;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Media.Imaging;
using Windows.UI.Xaml.Navigation;
using Service = DocumentTransitionUniversalApp.TransitionAppServices;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace DocumentTransitionUniversalApp.Views
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class WordSelectPartsPage : Page
    {
        #region Fields

        MainPage _source;
        public WordPartsPageData _pageData { get; private set; }

        #endregion

        #region Constructors

        public WordSelectPartsPage()
        {
            this.InitializeComponent();
            InitializeVariables();
        }

        #endregion

        #region public methods

        public void CopyDataToControl(WordPartsPageData data)
        {
            _pageData = data;
        }

        #endregion

        #region Private methods

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

        private void PrepareListOfItems(ObservableCollection<Service.PartsSelectionTreeElement> elements)
        {
            foreach (var element in elements)
            {
                TreeElementIcon icon = new TreeElementIcon(element.Type);
                var item = new PartsSelectionTreeElement(element.Id, element.ElementId, element.Type, element.Name, element.Indent, icon.GetIcon());
                foreach (var child in element.Childs)
                {
                    item.SetChild(PartsSelectionTreeElement.ConvertToPartsSelectionTreeElement(child));
                }

                _pageData.SelectionParts.Add(item);
            }
        }

        private void CreateSelectPartsUI(ObservableCollection<PartsSelectionTreeElement> elements)
        {
            WordSelectPartsItems.Items.Clear();
            LazyLoadingItems<PartsSelectionTreeElement> lazyItems = new LazyLoadingItems<PartsSelectionTreeElement>(elements, PartsScrollViewer);
            lazyItems.PropertyChanged += LazyItems_PropertyChanged;
            foreach (var element in lazyItems.Items)
            {
                CreateButtonBlock(element);
            }
        }

        private void LazyItems_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {

            var lazy = sender as LazyLoadingItems<PartsSelectionTreeElement>;
            if (!lazy.IsPullRefresh)
            {
                WordSelectPartsItems.Items.Clear();
                foreach (var item in lazy.Items)
                {
                    CreateButtonBlock(item);
                }
            }
        }

        private void CreateButtonBlock(PartsSelectionTreeElement element)
        {
            Button button = new Button();

            if (element.Selected)
                button.Background = new SolidColorBrush((Color)Application.Current.Resources["SystemAccentColor"]);
            else
                button.Background = new SolidColorBrush(Colors.WhiteSmoke);

            button.Name = element.Id;
            button.Margin = new Thickness(element.Indent * 20, 5, 0, 5);

            Grid grid = new Grid();
            TextBlock textBlock = new TextBlock();
            textBlock.Text = element.Name;
            textBlock.Margin = new Thickness(30, 0, 0, 0);
            textBlock.VerticalAlignment = VerticalAlignment.Center;
            Image image = new Image();
            image.Source = new BitmapImage(new Uri(element.Icon));
            image.Stretch = Stretch.Uniform;
            image.Width = 30;
            image.Height = 30;
            image.HorizontalAlignment = HorizontalAlignment.Left;
            image.Margin = new Thickness(-4, 0, 0, 0);
            grid.Children.Add(image);
            grid.Children.Add(textBlock);

            button.Content = grid;

            ToolTip toolTip = new ToolTip();
            toolTip.Content = element.Name;
            ToolTipService.SetToolTip(button, toolTip);

            if ((string)comboBox.SelectedItem != null)
            {
                var comboItem = Data_Structures.ComboBoxItem.GetComboBoxItemByName(_pageData.ComboItems, (string)comboBox.SelectedItem);
                if (element.CheckIfCanBeSelected(Data_Structures.ComboBoxItem.GetComboBoxItemByName(_pageData.ComboItems, comboItem.Name).Name) && comboItem.Id != WordPartsPageData.AllItemsId)
                {
                    button.Tapped += Button_Tapped;
                }
                else
                    button.Background = new SolidColorBrush(Colors.DimGray);
            }
            else
                button.Background = new SolidColorBrush(Colors.DimGray);

            WordSelectPartsItems.Items.Add(button);

            foreach (var child in element.Childs)
            {
                CreateButtonBlock(child);
            }
        }

        private void Button_Tapped(object sender, TappedRoutedEventArgs e)
        {
            var ownerName = Data_Structures.ComboBoxItem.GetComboBoxItemByName(_pageData.ComboItems, (string)comboBox.SelectedItem).Name;
            var button = sender as Button;
            var selectedElement = _pageData.SelectionParts.Traverse(sp => sp.Childs).Single(el => el.Id == button.Name);
            if (selectedElement.Selected)
            {
                selectedElement.SelectItem(string.Empty);
                ChangeButtonsColor(selectedElement, button, Colors.WhiteSmoke);
            }
            else
            {
                selectedElement.SelectItem(ownerName);
                ChangeButtonsColor(selectedElement, button, (Color)Application.Current.Resources["SystemAccentColor"]);
            }
        }

        private void ChangeButtonsColor(PartsSelectionTreeElement element, Button button, Color color)
        {
            button.Background = new SolidColorBrush(color);
            foreach (var child in element.Childs)
            {
                var childButton = WordSelectPartsItems.Items.First(b => (b as Button).Name == child.Id);
                ChangeButtonsColor(child, childButton as Button, color);
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

            var serviceClient = MainPage.Service.GetInstance();
            try
            {
                var result = await serviceClient.GetWordPartsAsync(_source.FileName, _source.documentBinary);
                PrepareListOfItems(result.Body.GetWordPartsResult);
            }
            catch (System.ServiceModel.CommunicationException ex)
            {
                MessageDialog dialog = new MessageDialog(ex.Message);
                await dialog.ShowAsync();
            }
        }

        private async void InitPresentation()
        {
            var serviceClient = MainPage.Service.GetInstance();
            try
            {
                var result = await serviceClient.GetPresentationPartsAsync(_source.FileName, _source.documentBinary);
                PrepareListOfItems(result.Body.GetPresentationPartsResult);
            }
            catch (System.ServiceModel.CommunicationException ex)
            {
                MessageDialog dialog = new MessageDialog(ex.Message);
                await dialog.ShowAsync();
            }
        }

        private async void InitExcel()
        {
            var serviceClient = MainPage.Service.GetInstance();
            try
            {
                var result = await serviceClient.GetExcelPartsAsync(_source.FileName, _source.documentBinary);
                PrepareListOfItems(result.Body.GetExcelPartsResult);
            }
            catch (System.ServiceModel.CommunicationException ex)
            {
                MessageDialog dialog = new MessageDialog(ex.Message);
                await dialog.ShowAsync();
            }
        }

        #endregion
    }
}
