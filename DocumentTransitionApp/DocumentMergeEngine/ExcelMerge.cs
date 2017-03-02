using DocumentMergeEngine.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SplitDescriptionObjects;

namespace DocumentMergeEngine
{
    public class ExcelMerge : DocumentMerge, IMerge
    {
        public byte[] Run(List<PersonFiles> files)
        {
            throw new NotImplementedException();
        }
    }
}
