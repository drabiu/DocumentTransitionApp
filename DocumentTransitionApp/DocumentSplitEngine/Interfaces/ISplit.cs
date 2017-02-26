using DocumentEditPartsEngine;
using SplitDescriptionObjects;
using System.Collections.Generic;
using System.IO;

namespace DocumentSplitEngine.Interfaces
{
    public interface ISplit
    {
        List<PersonFiles> SaveSplitDocument(Stream document);
        void OpenAndSearchDocument(Stream docFile, Stream xmlFile);      
    }
}
