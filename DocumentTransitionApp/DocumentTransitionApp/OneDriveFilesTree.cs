using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentTransitionApp
{
	public class OneDriveFilesTreeElement
	{
		enum ElementType
		{
			Folder,
			File
		}

		ElementType Type { public get; set; }
		OneDriveFilesTreeElement Child { public get; set; }
		string Name { public get; set; }
		int Indent { public get; set; }

		public OneDriveFilesTreeElement()
		{
		}
	}
}
