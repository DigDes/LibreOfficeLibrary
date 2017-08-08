using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace LibreOfficeLibrary
{
	public class RevisionManager
	{
		public void AcceptAllRevisions(string filePath)
		{
			if (!File.Exists(filePath))
				throw new ArgumentException("The file doesn't exist");

			var tempDirPath = Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar);
			var tempFilePath = Path.Combine(tempDirPath, Path.GetRandomFileName());
			File.Copy(filePath, tempFilePath);

			var worker = new LibreOfficeWorker();
			worker.DoWork($"macro:///LibreOfficeLibrary.Module1.AcceptAllChanges(\"{tempFilePath}\")");

			File.Delete(filePath);
			File.Move(tempFilePath, filePath);
		}
	}
}
