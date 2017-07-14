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
	public class DocumentConverter
	{
		private const int TimeForWaiting = 60;

		/// <summary>
		/// Convert document to PDF format
		/// </summary>
		/// <param name="filePath">The absolute path of the document that will be converted</param>
		/// <param name="targetPath">The absolute path of the target document for the convertion</param>
		public void ConvertToPdf(string filePath, string targetPath)
		{
			var directoryName = Path.GetDirectoryName(filePath) ?? string.Empty;
			var tempOutputFilePath = Path.Combine(directoryName, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

			var pathContainsSpaces = false;
			var oldPath = filePath;
			if (filePath.Contains(" "))
			{
				string newPath = Path.Combine(directoryName, Path.GetRandomFileName());
				File.Move(filePath, newPath);
				pathContainsSpaces = true;
				filePath = newPath;
			}

			filePath = GetRelativePath(directoryName, filePath);
			directoryName = GetRelativePath(Path.Combine(Directory.GetCurrentDirectory(), "temp.temp"), directoryName);

			var worker = new Thread(() => DoConvert(directoryName, filePath));
			worker.IsBackground = true;
			worker.Start();

			if (!worker.Join(TimeSpan.FromSeconds(TimeForWaiting)))
			{
				worker.Abort();

				return;
			}

			if (pathContainsSpaces)
			{
				tempOutputFilePath = Path.Combine(directoryName, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
				File.Move(filePath, oldPath);
			}

			if (File.Exists(tempOutputFilePath))
				File.Move(tempOutputFilePath, targetPath);
		}

		private void DoConvert(string directoryName, string filePath)
		{
			Process process = new Process
			{
				StartInfo = new ProcessStartInfo
				{
					WindowStyle = ProcessWindowStyle.Hidden,
					FileName = "soffice.exe",
					Arguments = $"/C -headless -writer -convert-to pdf -outdir {directoryName} {filePath}"
				}
			};
			process.Start();
			process.WaitForExit();
		}

		private static string GetRelativePath(string directoryPath, string filePath)
		{
			var absDirPath = Path.GetFullPath(directoryPath);
			var absFilePath = Path.GetFullPath(filePath);
			var directory = new Uri(absDirPath);
			var file = new Uri(absFilePath);
			return Uri.UnescapeDataString(directory.MakeRelativeUri(file)
				.ToString()
				.Replace('/', Path.DirectorySeparatorChar)
			);
		}
	}
}
