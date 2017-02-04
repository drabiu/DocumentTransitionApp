using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentEditPartsEngine
{
    public class PartsSelectionTreeElement
    {
        public string Id { get; set; }
        public string ElementId { get; set; }
        //public ElementType Type { get; private set; }
        public List<PartsSelectionTreeElement> Childs { get; set; }
        public string Name { get; set; }
        public int Indent { get; set; }
        public string OwnerName { get; set; }
        public bool Selected { get; set; }

        public PartsSelectionTreeElement()
        {
        }

        public PartsSelectionTreeElement(string id, string name, int indent)
        {
            this.Id = id;
            //this.Type = type;
            this.Name = name;
            this.Indent = indent;
            this.Childs = new List<PartsSelectionTreeElement>();
        }

        public PartsSelectionTreeElement(string id, string elementId, string name, int indent) : this(id, name, indent)
        {
            this.ElementId = elementId;
        }
    }
}
