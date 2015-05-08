using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SplitDescriptionObjects
{
	public class Parts
	{
		public Guid Guid { get; private set; }
		public Stream DocumentStream { get; private set; }

		public Parts(Stream stream)
		{
			this.Guid = Guid.NewGuid();
			this.DocumentStream = stream;
		}
	}
}
