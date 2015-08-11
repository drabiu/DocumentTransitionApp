using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentTransitionPhoneApp
{
	public class OneDriveFilesTreeElement
	{
		public enum ElementType
		{
			Folder,
			File
		}

		public ElementType Type { get; private set; }
		public OneDriveFilesTreeElement Child { get; private set; }
		public string Name { get; private set; }
		public int Indent { get; private set; }

		public OneDriveFilesTreeElement()
		{
		}
	}
}
