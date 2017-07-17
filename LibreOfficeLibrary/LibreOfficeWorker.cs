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
		private class Worker
		{
			public Worker(object parameterObject)
			{
				Process = new Process
				{
					StartInfo = new ProcessStartInfo
					{
						WindowStyle = ProcessWindowStyle.Hidden,
						FileName = "soffice.exe",
						Arguments = (string)parameterObject
					}
				};
			}
			public void DoWork()
			{
				using (Process)
				{
					Process.Start();
					Process.WaitForExit();
				}
			}
			public Process Process { get; }
		}
		public void DoWork(string argumentString)
		{
			var worker = new Worker(argumentString);
			var thread = new Thread(worker.DoWork);
			thread.IsBackground = true;
			thread.Start();

			if (!thread.Join(TimeSpan.FromSeconds(TimeForWaiting)))
			{
				worker.Process.Kill();
				worker.Process.Dispose();
				thread.Abort();
				throw new ConvertDocumentException("LibreOffice process didn't respond within the expected time");
			}
		}
	}
}
