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

        public IEnumerable<PartsSelectionTreeElement> CreatePartsSelectionTreeElements(OpenXmlElement element, PartsSelectionTreeElement parent, int id, Predicate<OpenXmlElement> supportedType, int indent, bool visible)
        {
            List<PartsSelectionTreeElement> result = new List<PartsSelectionTreeElement>();
            if (supportedType(element))
            {
                Index = id;
                PartsSelectionTreeElement elementToAdd = _documentParts.GetParagraphSelectionTreeElement(element, parent, ref Index, supportedType, indent, visible);

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
                        //if (supportedType(elmentChild) && !elmentChild is Run)
                        //    Index++;

                        if (parent != null)
                        {
                            var parentElement = elementToAdd ?? parent;
                            var childrenToAdd = CreatePartsSelectionTreeElements(elmentChild, parentElement, Index, supportedType, indent, visible);
                            foreach (var child in childrenToAdd)
                            {
                                child.Parent = parentElement;
                                parentElement.Childs.Add(child);
                            }

                        }
                        else
                        {
                            result.AddRange(CreatePartsSelectionTreeElements(elmentChild, elementToAdd, Index, supportedType, indent, visible));
                        }
                    }

                }
            }

            return result;
        }
    }
}
