using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Threading;

namespace LibreOfficeLibrary
{
	public class LibreOfficeWorker
	{
		private const int TimeForWaiting = 60;
		public void DoWork(string argumentString)
		{
			var worker = new Thread(DoWorkInThread);
			worker.IsBackground = true;
			worker.Start(argumentString);

			if (!worker.Join(TimeSpan.FromSeconds(TimeForWaiting)))
			{
				worker.Abort();
				throw new ConvertDocumentException("LibreOffice process didn't respond within the expected time");
			}
		}
		private void DoWorkInThread(object parameterObject)
		{
			var process = new Process
			{
				StartInfo = new ProcessStartInfo
				{
					WindowStyle = ProcessWindowStyle.Hidden,
					FileName = "soffice.exe",
					Arguments = (string)parameterObject
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
