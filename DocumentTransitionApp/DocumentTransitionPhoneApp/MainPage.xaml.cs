using System;
using System.IO;
using System.IO.IsolatedStorage;

using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using System.Dynamic;
using System.Threading.Tasks;

using Microsoft.Phone.Controls;
using Microsoft.Phone.Shell;
using Microsoft.Live;
using Microsoft.Live.Controls;

using Windows.Storage;

using DocumentTransitionPhoneApp.Resources;
using Service = DocumentTransitionPhoneApp.TransitionAppServices;

namespace DocumentTransitionPhoneApp
{
	public partial class MainPage : PhoneApplicationPage
	{
		private LiveConnectClient Client;
		private LiveAuthClient AuthClient;
		private IDictionary<string, OneDriveFilesTreeElement> FilesTreeElements;
		private SelectedFile DocumentFile;
		private SelectedFile SplitFile;

		// Constructor
		public MainPage()
		{
			InitializeComponent();
			FilesTreeElements = new Dictionary<string, OneDriveFilesTreeElement>();

			// Sample code to localize the ApplicationBar
			//BuildLocalizedApplicationBar();
		}

		private void loginButton_SessionChanged(object sender, LiveConnectSessionChangedEventArgs e)
		{
			//if (e != null && e.Status == LiveConnectSessionStatus.Connected)
			//{
			//	//the session status is connected so we need to set this session status to client
			//	this.Client = new LiveConnectClient(e.Session);
			//}
			//else
			//{
			//	this.Client = null;
			//}
		}

		public async Task<LiveConnectClient> Login()
		{
			AuthClient = new LiveAuthClient("000000004015B444");

			LiveLoginResult result = await AuthClient.InitializeAsync(new string[] { "wl.signin", "wl.skydrive" });
			if (result.Status == LiveConnectSessionStatus.Connected)
			{
				return new LiveConnectClient(result.Session);
			}
			result = await AuthClient.LoginAsync(new string[] { "wl.signin", "wl.skydrive" });
			if (result.Status == LiveConnectSessionStatus.Connected)
			{
				return new LiveConnectClient(result.Session);
			}
			return null;
		}

		public async void GetFolders(LiveConnectClient client)
		{
			try
			{
				LiveOperationResult operationResult = await client.GetAsync("me/skydrive/files");
				dynamic result = (operationResult.Result as dynamic).data;
				//OneDriveExplorerPanel.Children.Clear();
				await CreateDynamicFilesTree(client, result, 0);
				//CreateFileExploreUI();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}			
		}

		private async void LogOn_Click(object sender, RoutedEventArgs e)
		{
			Client = await Login();
			FilesTreeElements.Clear();
			GetFolders(Client);
		}

		private async Task CreateDynamicFilesTree(LiveConnectClient client, IList<object> listedItems, int indent)
		{
			foreach (var listItem in listedItems)
			{
				string name = (listItem as dynamic).name;

				if ((listItem as dynamic).type == "folder" || (listItem as dynamic).type == "album")
				{
					OneDriveFilesTreeElement newElement = new OneDriveFilesTreeElement((listItem as dynamic).id, OneDriveFilesTreeElement.ElementType.Folder, name, indent);
					AddElementToList((listItem as dynamic).parent_id, newElement);
					string uri = (listItem as dynamic).id + "/files";
					try
					{
						LiveOperationResult operationResult = await client.GetAsync(uri);
						dynamic result = (operationResult.Result as dynamic).data;
						CreateDynamicFilesTree(client, result, indent + 1);
						CreateFileExploreUI();
					}
					catch (Exception ex)
					{
						MessageBox.Show(ex.Message);
					}
				}
				else
				{
					OneDriveFilesTreeElement newElement = new OneDriveFilesTreeElement((listItem as dynamic).id, OneDriveFilesTreeElement.ElementType.File, name, indent);
					AddElementToList((listItem as dynamic).parent_id, newElement);
				}
			}
		}

		private void AddElementToList(string parentId, OneDriveFilesTreeElement element)
		{
			OneDriveFilesTreeElement parent;
			if (FilesTreeElements.TryGetValue(parentId, out parent))
			{
				parent.SetChild(element);
			}

			FilesTreeElements.Add(element.Id, element);
		}

		private void CreateFileExploreUI(Func<OneDriveFilesTreeElement, bool> filter)
		{
			OneDriveExplorerPanel.Children.Clear();
			foreach (KeyValuePair<string, OneDriveFilesTreeElement> element in FilesTreeElements)
			{
				if (element.Value.Indent == 0)
				{
					foreach (OneDriveFilesTreeElement child in element.Value.GetFilesTreeList().Where(filter))
					{
						CreateTextBlock(child);
					}
				}
			}
		}

		private void CreateFileExploreUI(string filterName)
		{
			CreateFileExploreUI(el => el.Name.Contains(filterName));
		}

		private void CreateFileExploreUI()
		{
			CreateFileExploreUI(el => true);
		}

		private void CreateTextBlock(OneDriveFilesTreeElement element)
		{
			TextBlock textBlock = new TextBlock();
			textBlock.Name = element.Id;
			textBlock.TextWrapping = TextWrapping.Wrap;
			textBlock.Margin = new Thickness(element.Indent * 20, 5, 0, 5);
			textBlock.Text = element.Name;
			if (element.Type == OneDriveFilesTreeElement.ElementType.Folder)
			{
				textBlock.FontWeight = FontWeights.Bold;
			}
			else if (element.Type == OneDriveFilesTreeElement.ElementType.File)
			{
				textBlock.Tap += textBlock_Tap;
			}

			OneDriveExplorerPanel.Children.Add(textBlock);
		}

		private async void textBlock_Tap(object sender, System.Windows.Input.GestureEventArgs e)
		{
			TextBlock block = (sender as TextBlock);
			if (block.Text.Contains(".docx") || block.Text.Contains(".xlsx") || block.Text.Contains(".pptx"))
			{
				DocumentFile = await GetFile(block.Text, block.Name);
				DocumentLabelTextBlock.Text = block.Text;
				RunSplitButton.IsEnabled = true;
			}
			else if (block.Text.Contains(".xml"))
			{
				SplitFile = await GetFile(block.Text, block.Name);
				SplitLabelTextBlock.Text = block.Text;
				RunMergeButton.IsEnabled = true;
			}
			else
			{
				MessageBox.Show("This is a wrong File.");
			}
		}

		private async Task<SelectedFile> GetFile(string fileName, string fileId)
		{
			string filePath = "shared/transfers/" + fileName;
			filePath = Uri.EscapeUriString(filePath);
			Uri fileUri = new Uri(filePath, UriKind.Relative);
			await Client.BackgroundDownloadAsync(fileId + "/Content", fileUri);

			System.IO.Stream readStream = new MemoryStream();
			string root = ApplicationData.Current.LocalFolder.Path;
			var storageFolder = await StorageFolder.GetFolderFromPathAsync(root + @"\shared\transfers");
            StorageFile storageFile = await storageFolder.GetFileAsync(Uri.EscapeUriString(fileName));
            if ( storageFile != null )
            {
				readStream = await storageFile.OpenStreamForReadAsync();
			}

			return new SelectedFile(readStream, fileName);
		}

		private void SearchButton_Click(object sender, RoutedEventArgs e)
		{
			CreateFileExploreUI(FilterTextBox.Text);
		}

		private void RunSplitButton_Click(object sender, RoutedEventArgs e)
		{
			Service.Service1SoapClient serviceClient = new Service.Service1SoapClient();
			serviceClient.SplitDocumentCompleted += serviceClient_SplitDocumentCompleted;

			byte[] fileBinary = new byte[DocumentFile.File.Length];
			DocumentFile.File.Read(fileBinary, 0, (int)DocumentFile.File.Length);

			byte[] splitBinary = new byte[SplitFile.File.Length];
			SplitFile.File.Read(splitBinary, 0, (int)SplitFile.File.Length);

			serviceClient.SplitDocumentAsync(DocumentFile.FileName, fileBinary, splitBinary);
		}

		void serviceClient_SplitDocumentCompleted(object sender, Service.SplitDocumentCompletedEventArgs e)
		{
			//IEnumerable<Service.PersonFiles> eee = e.Result.AsEnumerable();
			var eee = e.Result;
		}

		private void RunMergeButton_Click(object sender, RoutedEventArgs e)
		{

		}
	}
}