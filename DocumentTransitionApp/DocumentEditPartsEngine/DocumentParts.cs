﻿using DocumentEditPartsEngine.Interfaces;
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
                PartsSelectionTreeElement elementToAdd = _documentParts.GetParagraphSelectionTreeElement(element, parent, id, supportedType, indent, visible);

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
                            parentElement.Childs.AddRange(CreatePartsSelectionTreeElements(elmentChild, parentElement, Index, supportedType, indent, visible));
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
