using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentEditPartsEngine.Interfaces;
using System.IO;

namespace DocumentEditPartsEngine
{
    public class ExcelDocumentParts : IDocumentParts
    {
        public List<PartsSelectionTreeElement> Get(Stream file)
        {
            throw new NotImplementedException();
        }
    }
}
