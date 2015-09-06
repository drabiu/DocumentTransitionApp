using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace DocumentTransitionPhoneApp
{
	public class SelectedFile
	{
		public Stream File { get; private set; }
		public string FileName { get; private set;}

		public SelectedFile(Stream file, string fileName)
		{
			this.File = file;
			this.FileName = fileName;
		}
	}

	public class OneDriveFilesTreeElement
	{
		public enum ElementType
		{
			Folder,
			File
		}

		public string Id { get; private set; }
		public ElementType Type { get; private set; }
		public IList<OneDriveFilesTreeElement> Childs { get; private set; }
		public string Name { get; private set; }
		public int Indent { get; private set; }

		public OneDriveFilesTreeElement(string id, ElementType type, string name, int indent)
		{
			this.Id = id;
			this.Type = type;
			this.Name = name;
			this.Indent = indent;
			this.Childs = new List<OneDriveFilesTreeElement>();
		}

		public OneDriveFilesTreeElement(string id, ElementType type, OneDriveFilesTreeElement child, string name, int indent)
			: this(id, type, name, indent)
		{
			this.Childs.Add(child);
		}

		public void SetChild(OneDriveFilesTreeElement child)
		{
			this.Childs.Add(child);
		}

		public IList<OneDriveFilesTreeElement> GetFilesTreeList()
		{
			List<OneDriveFilesTreeElement> result = new List<OneDriveFilesTreeElement>();
			result.Add(this);

			foreach (OneDriveFilesTreeElement child in Childs)
			{
				result.AddRange(GetChilds(child));
			}

			return result;
		}

		public IList<OneDriveFilesTreeElement> GetChilds(OneDriveFilesTreeElement element)
		{
			List<OneDriveFilesTreeElement> result = new List<OneDriveFilesTreeElement>();

			result.Add(element);

			foreach (OneDriveFilesTreeElement child in element.Childs)
			{
				result.AddRange(GetChilds(child));
			}

			return result;
		}
	}
}
