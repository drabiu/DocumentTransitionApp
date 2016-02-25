using System.Collections.Generic;

namespace DocumentTransitionUniversalApp
{
	public class ElementTypes
	{
		public enum WordElementType
		{
			Paragraph,
			Table,
			Picture
		}

		public enum ExcelElementType
		{
			Sheet
		}

		public enum PresentationElementType
		{
			Slide
		}
	}

	public class PartsSelectionTreeElement<ElementType>
	{
		public string Id { get; private set; }
		public ElementType Type { get; private set; }
		public IList<PartsSelectionTreeElement<ElementType>> Childs { get; private set; }
		public string Name { get; private set; }
		public int Indent { get; private set; }
		public bool Selected { get; set; }
        private int _ownerId { get; set; }
        int _allItemsId = 0;

        public PartsSelectionTreeElement(string id, ElementType type, string name, int indent)
		{
			this.Id = id;
			this.Type = type;
			this.Name = name;
			this.Indent = indent;
			this.Childs = new List<PartsSelectionTreeElement<ElementType>>();
			this.Selected = false;
            this._ownerId = _allItemsId;
		}

		public PartsSelectionTreeElement(string id, ElementType type, PartsSelectionTreeElement<ElementType> child, string name, int indent)
			: this(id, type, name, indent)
		{
			this.Childs.Add(child);
		}

		public void SetChild(PartsSelectionTreeElement<ElementType> child)
		{
			this.Childs.Add(child);
		}

		public IList<PartsSelectionTreeElement<ElementType>> GetFilesTreeList()
		{
			List<PartsSelectionTreeElement<ElementType>> result = new List<PartsSelectionTreeElement<ElementType>>();
			result.Add(this);

			foreach (PartsSelectionTreeElement<ElementType> child in Childs)
			{
				result.AddRange(GetChilds(child));
			}

			return result;
		}

		public IList<PartsSelectionTreeElement<ElementType>> GetChilds(PartsSelectionTreeElement<ElementType> element)
		{
			List<PartsSelectionTreeElement<ElementType>> result = new List<PartsSelectionTreeElement<ElementType>>();

			result.Add(element);

			foreach (PartsSelectionTreeElement<ElementType> child in element.Childs)
			{
				result.AddRange(GetChilds(child));
			}

			return result;
		}

        public bool CheckIfCanBeSelected(int ownerId)
        {
            if ((this._ownerId != _allItemsId && this._ownerId == ownerId) || (this._ownerId == _allItemsId && ownerId != _allItemsId))
                return true;
            else
                return false;
        }

        public void SelectItem(int ownerId)
        {
            Selected = !Selected;
            this._ownerId = ownerId;
        }
	}
}
