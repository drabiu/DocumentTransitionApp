using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI;
using Windows.UI.Core;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
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
		MainPage Source;
		List<ComboBoxItem> ComboItems;
        Dictionary<int, PartsSelectionTreeElement<ElementTypes.WordElementType>> SelectionParts;

		public WordSelectPartsPage()
		{
			this.InitializeComponent();
			ComboItems = new List<ComboBoxItem>();
			ComboItems.Add(new ComboBoxItem() { Id = 0, Name = "All" });
			comboBox.ItemsSource = ComboItems.Select(cmb => cmb.Name);
		}

		protected override void OnNavigatedTo(NavigationEventArgs e)
		{
			if (e.Parameter is MainPage)
			{
				this.Source = e.Parameter as MainPage;
			}
			base.OnNavigatedTo(e);
		}

		private void BackButton_Click(object sender, RoutedEventArgs e)
		{
			this.Frame.Navigate(typeof(MainPage), Source);
		}

		private void CreateSelectPartsUI(System.Collections.ObjectModel.ObservableCollection<DocumentTransitionUniversalApp.TransitionAppServices.PartsSelectionTreeElement> elements)
		{
            WordSelectPartsItems.Items.Clear();
            foreach (var element in elements)
            {
                CreateButtonBlock(new PartsSelectionTreeElement<ElementTypes.WordElementType>(element.Id, ElementTypes.WordElementType.Paragraph, element.Name, element.Indent));
            }
        }

        private void CreateButtonBlock(PartsSelectionTreeElement<ElementTypes.WordElementType> element)
		{
            Button button = new Button();
            button.Background = new SolidColorBrush(Colors.Transparent);
            button.Name = element.Id;
            button.Margin = new Thickness(element.Indent * 20, 5, 0, 5);
            button.Content = element.Name;
            if (element.CanSelect)
                button.Tapped += Button_Tapped;

            WordSelectPartsItems.Items.Add(button);
        }

        private void Button_Tapped(object sender, TappedRoutedEventArgs e)
        {
            var button = sender as Button;
            button.Background = new SolidColorBrush(Colors.Aqua);
        }

        private async void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			var result = await serviceClient.GetPartsAsync(Source.FileName, Source.documentBinary);
            CreateSelectPartsUI(result.Body.GetPartsResult);
		}
    }

    public class ComboBoxItem
	{
		public string Name { get; set; }
		public int Id { get; set; }
	}
}
