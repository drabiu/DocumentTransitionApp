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
}
