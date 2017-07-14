using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace LibreOfficeLibrary
{
	public class DocumentComparer
	{
		/// <summary>
		/// Compare two documents with LibreOffice
		/// </summary>
		public void Compare(string filePath, string fileToComparePath, string targetFilePath)
		{
			if (!File.Exists(filePath) || !File.Exists(fileToComparePath))
				throw new ArgumentException("One of files doesn't exist");
			if (File.Exists(targetFilePath))
				throw new ArgumentException("The target file exists");

			var tempDirPath = Path.GetTempPath();
			var tempFilePath = Path.Combine(tempDirPath, Path.GetRandomFileName());
			var tempFileToComparePath = Path.Combine(tempDirPath, Path.GetRandomFileName());

			File.Copy(filePath, tempFilePath);
			File.Copy(tempFilePath, tempFileToComparePath);

			var worker = new LibreOfficeWorker();
			worker.DoWork($"macro:///LibreOfficeLibrary.Module1.CompareDocuments(\"{tempFilePath}\",\"{tempFileToComparePath}\")");

			File.Move(tempFilePath, targetFilePath);
		}
	}
}
