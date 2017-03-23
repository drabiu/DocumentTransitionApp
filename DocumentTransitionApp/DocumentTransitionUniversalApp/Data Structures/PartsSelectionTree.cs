using DocumentTransitionUniversalApp.Helpers;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace DocumentTransitionUniversalApp
{
    public class PartsSelectionTreeElement
    {
        #region Fields

        public string Id { get; private set; }
        public string ElementId { get; private set; }
        public TransitionAppServices.ElementType Type { get; private set; }
        public List<PartsSelectionTreeElement> Childs { get; private set; }
        public string Name { get; private set; }
        public string Icon { get; private set; }
        public int Indent { get; private set; }
        public bool Selected { get; set; }
        private string _ownerName { get; set; }

        #endregion

        #region Constructors

        public PartsSelectionTreeElement(string id, TransitionAppServices.ElementType type, string name, int indent)
        {
            this.Id = id;
            this.Type = type;
            this.Name = name;
            this.Indent = indent;
            this.Childs = new List<PartsSelectionTreeElement>();
            this.Selected = false;
            this._ownerName = string.Empty;
        }

        public PartsSelectionTreeElement(string id, string elementId, TransitionAppServices.ElementType type, string name, int indent, string icon) : this(id, type, name, indent)
        {
            this.ElementId = elementId;
            this.Icon = icon;
        }

        public PartsSelectionTreeElement(string id, string elementId, TransitionAppServices.ElementType type, string name, int indent) : this(id, type, name, indent)
        {
            this.ElementId = elementId;
        }

        public PartsSelectionTreeElement(string id, TransitionAppServices.ElementType type, PartsSelectionTreeElement child, string name, int indent)
            : this(id, type, name, indent)
        {
            this.Childs.Add(child);
        }

        public PartsSelectionTreeElement(string id, string elementId, TransitionAppServices.ElementType type, string name, int indent, bool selected)
            : this(id, elementId, type, name, indent)
        {
            this.Selected = selected;
        }

        public PartsSelectionTreeElement(string id, string elementId, TransitionAppServices.ElementType type, string name, int indent, bool selected, string ownerName)
            : this(id, elementId, type, name, indent, selected)
        {
            this._ownerName = ownerName;
        }

        #endregion

        #region Public methods

        public void SetChild(PartsSelectionTreeElement child)
        {
            this.Childs.Add(child);
        }

        public void SetChildRecursive()
        {

        }

        public void AddChilds(List<PartsSelectionTreeElement> childs)
        {
            foreach (var child in childs)
            {
                SetChild(child);
            }
        }

        public void AddChildsRecursive(List<PartsSelectionTreeElement> childs)
        {
            AddChilds(childs);
            foreach (var child in childs)
            {
                AddChildsRecursive(child.Childs);
            }
        }

        public IList<PartsSelectionTreeElement> GetFilesTreeList()
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            result.Add(this);

            foreach (PartsSelectionTreeElement child in Childs)
            {
                result.AddRange(GetChilds(child));
            }

            return result;
        }

        public IList<PartsSelectionTreeElement> GetChilds(PartsSelectionTreeElement element)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();

            result.Add(element);

            foreach (PartsSelectionTreeElement child in element.Childs)
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
            foreach (var child in Childs)
            {
                if (child.Selected != Selected)
                    child.SelectItem(ownerName);
            }
        }

        public TransitionAppServices.PartsSelectionTreeElement ConvertToServicePartsSelectionTreeElement()
        {
            var part = new TransitionAppServices.PartsSelectionTreeElement();
            part.Id = this.Id;
            part.ElementId = this.ElementId;
            part.Indent = this.Indent;
            part.Name = this.Name;
            part.OwnerName = this._ownerName;
            part.Selected = this.Selected;
            part.Type = this.Type;
            part.Childs = new ObservableCollection<TransitionAppServices.PartsSelectionTreeElement>();
            foreach (var child in this.Childs)
            {
                part.Childs.Add(child.ConvertToServicePartsSelectionTreeElement());
            }

            return part;
        }

        #endregion

        #region Public static methods

        public static PartsSelectionTreeElement ConvertToPartsSelectionTreeElement(TransitionAppServices.PartsSelectionTreeElement element)
        {
            TreeElementIcon icon = new TreeElementIcon(element.Type);
            var part = new PartsSelectionTreeElement(element.Id, element.ElementId, element.Type, element.Name, element.Indent, icon.GetIcon());
            part._ownerName = element.OwnerName;
            part.Selected = element.Selected;
            foreach (var child in element.Childs)
            {
                part.Childs.Add(ConvertToPartsSelectionTreeElement(child));
            }

            return part;
        }

        #endregion

        #region Private methods


        #endregion
    }
}
