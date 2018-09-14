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
		/// <summary>
		/// Convert document to PDF format
		/// </summary>
		public void ConvertToPdf(string filePath, string targetPath, string profileLocation = null)
		{
			if (!File.Exists(filePath))
				throw new ArgumentException("The file doesn't exist");
			if (File.Exists(targetPath))
				throw new ArgumentException("The target file exists");

			if (profileLocation == null)
			{
				var path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
				var versionDir = Directory.GetDirectories(Path.Combine(path, "libreoffice")).FirstOrDefault();
				if(versionDir==null)
					return;

				profileLocation = versionDir;
			}

			var profilePath = Path.GetFullPath(profileLocation).Replace('\\', '/').Replace(" ", "%20");
			var tempDirPath = Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar);
			var tempFilePath = Path.Combine(tempDirPath, Path.GetRandomFileName());
			var tempOutputFilePath = Path.Combine(tempDirPath, Path.GetFileNameWithoutExtension(tempFilePath) + ".pdf");

			File.Copy(filePath, tempFilePath);

			var worker = new LibreOfficeWorker();
			worker.DoWork($"/C -headless -writer -convert-to pdf -outdir \"{tempDirPath}\" \"{tempFilePath}\" \"-env:UserInstallation=file:///{profilePath}/\"");

			if (File.Exists(tempOutputFilePath))
				File.Move(tempOutputFilePath, targetPath);
		}
	}
}
