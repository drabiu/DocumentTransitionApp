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
using Service = DocumentTransitionUniversalApp.TransitionAppServices;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace DocumentTransitionUniversalApp.Views
{
	/// <summary>
	/// An empty page that can be used on its own or navigated to within a Frame.
	/// </summary>
	public sealed partial class WordSelectPartsPage : Page
	{
		MainPage _source;
		List<ComboBoxItem> _comboItems;
        int _lastId;
        int _allItemsId = 0;
        List<PartsSelectionTreeElement<ElementTypes.WordElementType>> _selectionParts;

		public WordSelectPartsPage()
		{
			this.InitializeComponent();
            InitializeVariables();
        }

		protected override void OnNavigatedTo(NavigationEventArgs e)
		{
			if (e.Parameter is MainPage)
			{
				this._source = e.Parameter as MainPage;
			}

            InitializeItems();
            base.OnNavigatedTo(e);
		}

		private void BackButton_Click(object sender, RoutedEventArgs e)
		{
			this.Frame.Navigate(typeof(MainPage), _source);
		}

        private void PrepareListOfItems(System.Collections.ObjectModel.ObservableCollection<DocumentTransitionUniversalApp.TransitionAppServices.PartsSelectionTreeElement> elements)
        {
            foreach (var element in elements)
            {
                var item = new PartsSelectionTreeElement<ElementTypes.WordElementType>(element.Id, ElementTypes.WordElementType.Paragraph, element.Name, element.Indent);
                _selectionParts.Add(item);
            }
        }

		private void CreateSelectPartsUI(IEnumerable<PartsSelectionTreeElement<ElementTypes.WordElementType>> elements)
		{
            WordSelectPartsItems.Items.Clear();
            foreach (var element in elements)
            {
                CreateButtonBlock(element);
            }
        }

        private void CreateButtonBlock(PartsSelectionTreeElement<ElementTypes.WordElementType> element)
		{
            Button button = new Button();

            if (element.Selected)
                button.Background = new SolidColorBrush(Colors.Honeydew);
            else
                button.Background = new SolidColorBrush(Colors.Transparent);

            button.Name = element.Id;
            button.Margin = new Thickness(element.Indent * 20, 5, 0, 5);
            button.Content = element.Name;

            if ((string)comboBox.SelectedItem != null && element.CheckIfCanBeSelected(ComboBoxItem.GetComboBoxItemByName(_comboItems, (string)comboBox.SelectedItem).Id))
                button.Tapped += Button_Tapped;
            else
                button.Background = new SolidColorBrush(Colors.DimGray);

            WordSelectPartsItems.Items.Add(button);
        }

        private void Button_Tapped(object sender, TappedRoutedEventArgs e)
        {
            var ownerId = ComboBoxItem.GetComboBoxItemByName(_comboItems, (string)comboBox.SelectedItem).Id;
            var button = sender as Button;
            button.Background = new SolidColorBrush(Colors.Honeydew);
            var selectedElement = _selectionParts.Single(el => el.Id == button.Name);
            selectedElement.SelectItem(ownerId);
        }

        private async void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
            CreateSelectPartsUI(_selectionParts);
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
            if (_comboItems.Any(it => it.Name == name))
            {
                var dialog = new MessageDialog("There is already a person with this name");
                await dialog.ShowAsync();
            }
            else
            {
                _comboItems.Add(new ComboBoxItem() { Id = ++_lastId, Name = name });
            }

            comboBox.ItemsSource = _comboItems.Select(cb => cb.Name);
        }

        private void InitializeVariables()
        {
            _comboItems = new List<ComboBoxItem>();
            _comboItems.Add(new ComboBoxItem() { Id = _lastId = _allItemsId, Name = "All" });
            comboBox.ItemsSource = _comboItems.Select(cb => cb.Name);
            _selectionParts = new List<PartsSelectionTreeElement<ElementTypes.WordElementType>>();
            AddButton.IsEnabled = false;
        }

        private async void InitializeItems()
        {
            Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
            var result = await serviceClient.GetPartsAsync(_source.FileName, _source.documentBinary);
            PrepareListOfItems(result.Body.GetPartsResult);
        }
    }

    public class ComboBoxItem
	{
		public string Name { get; set; }
		public int Id { get; set; }

        public static ComboBoxItem GetComboBoxItemByName(IEnumerable<ComboBoxItem> items, string name)
        {
            return items.Single(it => it.Name == name);
        }
    }
}
