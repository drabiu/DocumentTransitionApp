﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentTransitionApp
{
	public class OneDriveFilesTreeElement
	{
		enum ElementType
		{
			Folder,
			File
		}

		ElementType Type { public get; set; }
	}
}
