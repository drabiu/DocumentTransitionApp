﻿using System;
using System.IO;
using System.IO.IsolatedStorage;

using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using System.Dynamic;

using Microsoft.Phone.Controls;
using Microsoft.Phone.Shell;
using Microsoft.Live;
using Microsoft.Live.Controls;

using DocumentTransitionPhoneApp.Resources;
using System.Threading.Tasks;

namespace DocumentTransitionPhoneApp
{
	public partial class MainPage : PhoneApplicationPage
	{
		private LiveConnectClient Client;
		private LiveAuthClient AuthClient;
		private IDictionary<string, OneDriveFilesTreeElement> FilesTreeElements;

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
					OneDriveFilesTreeElement newElement = new OneDriveFilesTreeElement(OneDriveFilesTreeElement.ElementType.Folder, name, indent);
					AddElementToList((listItem as dynamic).id, (listItem as dynamic).parent_id, newElement);
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
					OneDriveFilesTreeElement newElement = new OneDriveFilesTreeElement(OneDriveFilesTreeElement.ElementType.File, name, indent);
					AddElementToList((listItem as dynamic).id, (listItem as dynamic).parent_id, newElement);
				}
			}
		}

		private void AddElementToList(string id, string parentId, OneDriveFilesTreeElement element)
		{
			OneDriveFilesTreeElement parent;
			if (FilesTreeElements.TryGetValue(parentId, out parent))
			{
				parent.SetChild(element);

			}

			FilesTreeElements.Add(id, element);
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

		private void textBlock_Tap(object sender, System.Windows.Input.GestureEventArgs e)
		{
			TextBlock block = (sender as TextBlock);
			if (block.Text.Contains(".docx") || block.Text.Contains(".xlsx") || block.Text.Contains(".pptx"))
			{
				//do this
			}
			else
			{
				MessageBox.Show("This is a wrong File.");
			}
		}

		private void SearchButton_Click(object sender, RoutedEventArgs e)
		{
			CreateFileExploreUI(FilterTextBox.Text);
		}
	}
}