using DocumentEditPartsEngine.Interfaces;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;

namespace DocumentEditPartsEngine
{
    public class DocumentParts
    {
        IDocumentParts _documentParts;
        public int Index = 0;

        public DocumentParts(IDocumentParts documentParts)
        {
            _documentParts = documentParts;
        }

        public IEnumerable<PartsSelectionTreeElement> CreatePartsSelectionTreeElements(OpenXmlElement element, PartsSelectionTreeElement parent, int id, Predicate<OpenXmlElement> supportedType, int indent)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            if (supportedType(element))
            {
                PartsSelectionTreeElement elementToAdd = _documentParts.GetParagraphSelectionTreeElement(element, parent, id, supportedType, indent);

                if (elementToAdd != null)
                {
                    result.Add(elementToAdd);
                    indent--;
                }

                if (element.HasChildren)
                {
                    indent++;
                    foreach (var elmentChild in element.ChildElements)
                    {
                        Index++;
                        if (parent != null)
                        {
                            var parentElement = elementToAdd ?? parent;
                            parentElement.Childs.AddRange(CreatePartsSelectionTreeElements(elmentChild, parentElement, Index, supportedType, indent));
                        }
                        else
                        {
                            result.AddRange(CreatePartsSelectionTreeElements(elmentChild, elementToAdd, Index, supportedType, indent));
                        }
                    }

                }
            }

            return result;
        }
    }
}
