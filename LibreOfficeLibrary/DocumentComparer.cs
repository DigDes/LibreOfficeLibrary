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
		public void Compare(string filePath, string fileToComparePath, string targetFilePath)
		{
			if (!File.Exists(filePath) || !File.Exists(fileToComparePath))
				throw new ArgumentException("One of files doesn't exist");
			if (File.Exists(targetFilePath))
				throw new ArgumentException("The target file exists");

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
				throw new ConvertDocumentException("LibreOffice process didn't respond within the expected time");
			}

			File.Move(tempFilePath, targetFilePath);
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
			try
			{
				using (process)
				{
					process.Start();
					process.WaitForExit();
				}
			}
			finally
			{
				process.Kill();
			}
		}
	}
}
