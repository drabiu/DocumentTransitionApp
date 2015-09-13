using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
			EnableMergeButton();
		}

		private void buttonMerge_Click(object sender, RoutedEventArgs e)
		{

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
