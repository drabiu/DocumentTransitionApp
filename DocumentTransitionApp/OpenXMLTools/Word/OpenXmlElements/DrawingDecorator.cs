﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Interfaces;
using System.Text;
using P = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXMLTools.Word.OpenXmlElements
{
    public class DrawingDecorator : Drawing, IOpenXmlElementExtended
    {
        protected Drawing _drawing;

        public DrawingDecorator(OpenXmlElement drawing)
        {
            _drawing = drawing as Drawing;
        }

        public Drawing GetDrawing()
        {
            return _drawing;
        }

        public string GetElementName(int nameLength)
        {
            StringBuilder result = new StringBuilder();
            StringBuilder text = new StringBuilder();
            if (_drawing.Inline?.Graphic?.GraphicData != null)
            {
                foreach (var picture in _drawing.Inline?.Graphic.GraphicData.Elements<P.Picture>())
                {
                    foreach (var picProp in picture.Elements<P.NonVisualPictureProperties>())
                        text.Append(picProp.NonVisualDrawingProperties.Name);
                }
            }

            result = WordTools.GetWordsFromTextElements(text, nameLength);

            return result.ToString();
        }
    }
}
