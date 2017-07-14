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
		private const int TimeForWaiting = 60;

		/// <summary>
		/// Compare two documents with LibreOffice
		/// </summary>
		/// <param name="filePath">The path of the document that will be compared</param>
		/// <param name="fileToComparePath">The path of the document with which the document is compared</param>
		/// <param name="destinationFilePath">The path of the target document for the comparison</param>
		public void Compare(string filePath, string fileToComparePath, string destinationFilePath)
		{
			string tempDirPath = Path.GetTempPath();
			var tempFilePath = Path.Combine(tempDirPath, Path.GetRandomFileName());
			var tempFileToComparePath = Path.Combine(tempDirPath, Path.GetRandomFileName());

			File.Copy(filePath, tempFilePath);
			File.Copy(tempFilePath, tempFileToComparePath);

			var worker = new Thread(() => DoCompare(tempFilePath, tempFileToComparePath));
			worker.IsBackground = true;
			worker.Start();

			if (!worker.Join(TimeSpan.FromSeconds(TimeForWaiting)))
			{
				worker.Abort();
			}

			File.Move(tempFilePath, destinationFilePath);
		}

		private void DoCompare(string filePath, string fileToComparePath)
		{
			var process = new Process
			{
				StartInfo = new ProcessStartInfo
				{
					WindowStyle = ProcessWindowStyle.Hidden,
					FileName = "soffice.exe",
					Arguments = $"macro:///LibreOfficeLibrary.Module1.CompareDocuments(\"{filePath}\",\"{fileToComparePath}\")"
				}
			};
			process.Start();
			process.WaitForExit();
		}
	}
}
