using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentEditPartsEngine.Interfaces;

namespace DocumentEditPartsEngine
{
    public class DocumentPartsBuilder
    {
        public static IDocumentParts Build(string fileExtension)
        {
            IDocumentParts result;
            switch (fileExtension)
            {
                case (".docx"):
                    result = new WordDocumentParts();
                    break;
                case (".xlsx"):
                    result = new ExcelDocumentParts();
                    break;
                case (".pptx"):
                    result = new PresentationDocumentParts();
                    break;
                default:
                    result = new WordDocumentParts();
                    break;
            }

            return result;
        }
    }
}
