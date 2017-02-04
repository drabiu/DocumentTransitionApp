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
        public string ElementId { get; private set; }
		public ElementType Type { get; private set; }
		public IList<PartsSelectionTreeElement<ElementType>> Childs { get; private set; }
		public string Name { get; private set; }
		public int Indent { get; private set; }
		public bool Selected { get; set; }
        private string _ownerName { get; set; }

        public PartsSelectionTreeElement(string id, ElementType type, string name, int indent)
		{
			this.Id = id;
			this.Type = type;
			this.Name = name;
			this.Indent = indent;
			this.Childs = new List<PartsSelectionTreeElement<ElementType>>();
			this.Selected = false;
            this._ownerName = string.Empty;
		}

        public PartsSelectionTreeElement(string id, string elementId, ElementType type, string name, int indent) : this (id, type, name, indent)
        {
            this.ElementId = elementId;
        }

        public PartsSelectionTreeElement(string id, ElementType type, PartsSelectionTreeElement<ElementType> child, string name, int indent)
			: this(id, type, name, indent)
		{
			this.Childs.Add(child);
		}

        public PartsSelectionTreeElement(string id, string elementId, ElementType type, string name, int indent, bool selected) 
            : this (id, elementId, type, name, indent)
        {
            this.Selected = selected;
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

        public bool CheckIfCanBeSelected(string ownerName)
        {
            if ((!string.IsNullOrEmpty(this._ownerName) && this._ownerName == ownerName) || (string.IsNullOrEmpty(this._ownerName) && !string.IsNullOrEmpty(ownerName)))
                return true;
            else
                return false;
        }

        public void SelectItem(string ownerName)
        {
            Selected = !Selected;
            this._ownerName = ownerName;
        }

        public DocumentTransitionUniversalApp.TransitionAppServices.PartsSelectionTreeElement ConvertToPartsSelectionTreeElement()
        {
            var part = new TransitionAppServices.PartsSelectionTreeElement();
            part.Id = this.Id;
            part.ElementId = this.ElementId;
            part.Indent = this.Indent;
            part.Name = this.Name;
            part.OwnerName = this._ownerName;
            part.Selected = this.Selected;

            return part;
        }
    }
}
