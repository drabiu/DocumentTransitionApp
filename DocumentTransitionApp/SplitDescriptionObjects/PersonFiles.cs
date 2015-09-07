using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SplitDescriptionObjects
{
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
}
