using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace DocumentTransitionPhoneApp
{
	public interface ITransitionAppService
	{
		byte[] SplitDocument(string docName, byte[] docxFile, byte[] xmlFile);
		byte[] MergeDocument(PersonFiles[] files);
	}

	public class PersonFiles
	{
		public class FileData
		{
			public string Name { get; set; }
			public byte[] Data { get; set; }
		}

		public string Person { get; set; }
		public List<FileData> Files { get; set; }

		public PersonFiles()
		{
			Files = new List<FileData>();
		}
	}

	public class TransitionAppServiceClient : ITransitionAppService
	{
		private ITransitionAppService service;

		public TransitionAppServiceClient(Uri uri)
		{
			EndpointAddress adress = new EndpointAddress(uri);
			BasicHttpBinding binding = new BasicHttpBinding();
			binding.MaxBufferSize = 2000000;
			binding.MaxReceivedMessageSize = 2000000;

			ChannelFactory<ITransitionAppService> factory = new ChannelFactory<ITransitionAppService>(binding, adress);
			//service = factory.
		}

		public byte[] SplitDocument(string docName, byte[] docxFile, byte[] xmlFile)
		{
			throw new NotImplementedException();
		}

		public byte[] MergeDocument(PersonFiles[] files)
		{
			throw new NotImplementedException();
		}
	}
}
