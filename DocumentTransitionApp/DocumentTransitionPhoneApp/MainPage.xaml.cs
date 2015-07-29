using System;
using System.IO;
using System.IO.IsolatedStorage;

using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

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

		// Constructor
		public MainPage()
		{
			InitializeComponent();

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
			var result = await client.GetAsync(string.Format("{0}/files?limit=10", ""));
		}

		private async void Upload_Click(object sender, RoutedEventArgs e)
        {
			if (Client != null)
            {
                try
                {
                    string fileName = "sample.txt";
                    IsolatedStorageFile myIsolatedStorage = 
                    IsolatedStorageFile.GetUserStoreForApplication();//deletes the file if it already exists
                    if (myIsolatedStorage.FileExists(fileName))
                    {
                        myIsolatedStorage.DeleteFile(fileName);
                    }//now we use a StreamWriter to write inputBox.Text to the file and save it to IsolatedStorage
                    using (StreamWriter writeFile = new StreamWriter
                    (new IsolatedStorageFileStream(fileName, FileMode.Create, FileAccess.Write, myIsolatedStorage)))
                    {
                        writeFile.WriteLine("Hello world");
                        writeFile.Close();
                    }
                    IsolatedStorageFileStream isfs = myIsolatedStorage.OpenFile(fileName, FileMode.Open, FileAccess.Read);
                    var res = await Client.UploadAsync("me/skydrive", fileName, isfs, OverwriteOption.Overwrite);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please sign in with your Microsoft Account.");
            }
        }

		private async void LogOn_Click(object sender, RoutedEventArgs e)
		{
			Client = await Login();
		} 
	}
}