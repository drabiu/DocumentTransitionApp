using SplitDescriptionObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentMergeEngine.Interfaces
{
    public interface IMerge
    {
        byte[] Run(List<PersonFiles> files);
    }
}
