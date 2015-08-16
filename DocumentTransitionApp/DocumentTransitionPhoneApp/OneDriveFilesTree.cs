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
		public IList<OneDriveFilesTreeElement> Childs { get; private set; }
		public string Name { get; private set; }
		public int Indent { get; private set; }

		public OneDriveFilesTreeElement(ElementType type, string name, int indent)
		{
			this.Type = type;
			this.Name = name;
			this.Indent = indent;
			this.Childs = new List<OneDriveFilesTreeElement>();
		}

		public OneDriveFilesTreeElement(ElementType type, OneDriveFilesTreeElement child, string name, int indent)
			: this(type, name, indent)
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
			//foreach (OneDriveFilesTreeElement child in Childs)
			//{
			//	result.Add(child);
			//}

			//foreach (OneDriveFilesTreeElement child in Childs)
			//{
			//	result.AddRange(GetChilds(child));
			//}

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
			//foreach (OneDriveFilesTreeElement child in element.Childs)
			//{
			//	result.Add(child);
			//}

			//foreach (OneDriveFilesTreeElement child in element.Childs)
			//{
			//	result.AddRange(GetChilds(child));
			//}

			result.Add(element);

			foreach (OneDriveFilesTreeElement child in element.Childs)
			{
				result.AddRange(GetChilds(child));
			}

			return result;
		}
	}
}
