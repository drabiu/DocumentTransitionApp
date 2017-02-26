using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentSplitEngine
{
    public abstract class MarkerMapper
    {
        protected Split Xml { get; set; }
        protected string[] SubdividedParagraphs { get; set; }
    }
}
