using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
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

		private void CreateSelectPartsUI(IList<PartsSelectionTreeElement<ElementTypes.WordElementType>> elements)
		{
			PartsStackPanel.Children.Clear();
        }

		private async void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			var result = await serviceClient.GetPartsAsync(Source.FileName, Source.docxBinary);
		}
	}

	public class ComboBoxItem
	{
		public string Name { get; set; }
		public int Id { get; set; }
	}
}
