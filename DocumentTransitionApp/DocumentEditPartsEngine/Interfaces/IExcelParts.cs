using System.Collections.Generic;
using System.IO;

namespace DocumentEditPartsEngine.Interfaces
{
    public interface IExcelParts
    {
        List<PartsSelectionTreeElement> Get(Stream file);
    }
}
