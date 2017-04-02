using DocumentEditPartsEngine;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using SplitDescriptionObjects.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SplitDescriptionObjects
{
    public abstract class WordMarker : IWordMarker
    {
        protected Body DocumentBody;
        protected List<OpenXmlElement> ElementsList;

        public WordMarker(Body body)
        {
            DocumentBody = body;
            ElementsList = DocumentBody.ChildElements.ToList();
        }

        public int FindElement(string id)
        {
            throw new NotImplementedException();
        }

        protected static void SelectChildParts(IEnumerable<PartsSelectionTreeElement> parts, Person person)
        {
            foreach (var part in parts)
            {
                foreach (var child in part.Childs)
                {
                    SelectChildParts(child.Childs, person);
                    child.OwnerName = person.Email;
                    child.Selected = true;
                }
            }
        }
    }
}
