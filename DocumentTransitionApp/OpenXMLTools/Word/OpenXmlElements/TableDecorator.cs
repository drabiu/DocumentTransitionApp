using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXMLTools.Helpers;
using OpenXMLTools.Interfaces;
using System;
using System.Text;

namespace OpenXMLTools.Word.OpenXmlElements
{
    public class TableDecorator : Table, IOpenXmlElementExtended
    {
        private Table _table;

        public TableDecorator(OpenXmlElement table)
        {
            _table = table as Table;
        }

        public Table GetTable()
        {
            return _table;
        }

        public string GetElementName(int nameLength)
        {
            StringBuilder result = new StringBuilder();
            StringBuilder text = new StringBuilder(_table.InnerText);

            result = WordTools.GetWordsFromTextElements(text, nameLength);

            return result.ToString();
        }

        public ElementType GetElementType()
        {
            return ElementType.Table;
        }
    }
}
