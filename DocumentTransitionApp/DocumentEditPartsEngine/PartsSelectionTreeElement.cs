using OpenXMLTools.Helpers;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace DocumentEditPartsEngine
{
    [DataContract(IsReference = true)]
    public class PartsSelectionTreeElement
    {
        #region Fields

        [DataMember]
        public string Id { get; set; }
        [DataMember]
        public string ElementId { get; set; }
        [DataMember]
        public ElementType Type { get; set; }
        [DataMember]
        public PartsSelectionTreeElement Parent { get; set; }
        [DataMember]
        public List<PartsSelectionTreeElement> Childs { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public int Indent { get; set; }
        [DataMember]
        public string OwnerName { get; set; }
        [DataMember]
        public bool Selected { get; set; }
        [DataMember]
        public bool Visible { get; set; }

        #endregion

        #region Constructors

        public PartsSelectionTreeElement()
        {
            Visible = true;
        }

        public PartsSelectionTreeElement(string id, string name, int indent) : this()
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

        public bool IsListElement()
        {
            return Type == ElementType.BulletList || Type == ElementType.NumberedList;
        }

        #endregion
    }
}
