﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Services;

using DocumentSplitEngine;

namespace DocumentTransitionAppServices
{
	/// <summary>
	/// Summary description for Service1
	/// </summary>
	[WebService(Namespace = "http://tempuri.org/")]
	[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
	[System.ComponentModel.ToolboxItem(false)]
	// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
	// [System.Web.Script.Services.ScriptService]
	public class Service1 : System.Web.Services.WebService
	{
		[WebMethod]
		public Stream SplitDocument(string docName, Stream docxFile, Stream xmlFile)
		{
			ISplit run = new DocumentSplit(docName);
			run.OpenAndSearchWordDocument(docxFile, xmlFile);
			return run.SaveSplitDocument();
		}
	}
}