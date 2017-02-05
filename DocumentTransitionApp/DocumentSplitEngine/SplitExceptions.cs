using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentSplitEngine
{
	public class ElementToPersonPairException : Exception
	{
		public ElementToPersonPairException()
		{
		}

		public ElementToPersonPairException(string message)
			: base(message)
		{
		}

		public ElementToPersonPairException(string message, Exception inner)
			: base(message, inner)
		{
		}
	}
    
    public class SplitNameDifferenceExcception : Exception
    {
        private string _message;

        public override string Message
        {
            get
            {
                return string.Format("Document split service error: {0}", _message);
            }
        }

        public SplitNameDifferenceExcception(string message)
			: base(message)
		{
            _message = message;
        }
    } 	
}
