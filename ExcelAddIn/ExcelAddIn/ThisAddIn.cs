using Aras.IOM;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
	public partial class ThisAddIn
	{
		private const string userName = "";
		private const string password = "";
		private const string database = "";
		private const string innovatorServerUrl = "";
		private HttpServerConnection serverConnection;
		public Innovator Innovator { get; private set; }


		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			try
			{
				string passwordHash = Innovator.ScalcMD5(password);
				serverConnection = IomFactory.CreateHttpServerConnection(innovatorServerUrl, database, userName, passwordHash);
				var autorizationItem = (serverConnection as HttpServerConnection).Login();

				if (autorizationItem.isError())
				{
					serverConnection.Logout();
					throw new Exception(autorizationItem.getErrorString());
				}

				Innovator = autorizationItem.getInnovator();
			}
			catch(Exception exc)
			{
				MessageBox.Show(exc.Message);
			}
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
			serverConnection.Logout();
		}

		public Excel.Worksheet GetActiveWorkSheet()
		{
			return (Excel.Worksheet)Application.ActiveSheet;
		}

		public Excel.Workbooks GetActiveWorkbook()
		{
			return (Excel.Workbooks)Application.ActiveWorkbook;
		}

		public Excel.Application GetActiveApplication()
		{
			return (Excel.Application)Application.Application;
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new EventHandler(ThisAddIn_Startup);
			this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
