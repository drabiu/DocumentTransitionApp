using OpenXMLTools.Helpers;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace DocumentEditPartsEngine
{
    [Serializable]
    public class PartsSelectionTreeElement
    {
        #region Fields

        public string Id { get; set; }
        public string ElementId { get; set; }
        public ElementType Type { get; set; }
        public List<PartsSelectionTreeElement> Childs { get; set; }
        public string Name { get; set; }
        public int Indent { get; set; }
        public string OwnerName { get; set; }
        public bool Selected { get; set; }

        #endregion

        #region Constructors

        public PartsSelectionTreeElement()
        {
        }

        public PartsSelectionTreeElement(string id, string name, int indent)
        {
            this.Id = id;
            this.Name = name;
            this.Indent = indent;
            this.Childs = new List<PartsSelectionTreeElement>();
        }

        public PartsSelectionTreeElement(string id, string elementId, string name, int indent, ElementType type) : this(id, elementId, name, indent)
        {
            this.Type = type;
        }

        public PartsSelectionTreeElement(string id, string elementId, string name, int indent) : this(id, name, indent)
        {
            this.ElementId = elementId;
        }

        public PartsSelectionTreeElement(string id, string name, int indent, ElementType type) : this(id, name, indent)
        {
            this.Type = type;
        }

        #endregion

        #region Public methods

        public void SetChild(PartsSelectionTreeElement child)
        {
            this.Childs.Add(child);
        }

        #endregion
    }
}
