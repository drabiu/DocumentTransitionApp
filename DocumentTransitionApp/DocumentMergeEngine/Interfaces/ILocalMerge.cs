using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentMergeEngine.Interfaces
{
    [Obsolete]
    public interface ILocalMerge
    {
        void Run(string path);
    }

}
