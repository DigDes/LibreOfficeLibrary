using System;
using System.Collections.Generic;
using System.Text;

namespace LibreOfficeLibrary
{
	public class ConvertDocumentException : Exception
	{
		public ConvertDocumentException()
		{
		}

		public ConvertDocumentException(string message)
			: base(message)
		{
		}

		public ConvertDocumentException(string message, Exception inner)
			: base(message, inner)
		{
		}
	}
}
