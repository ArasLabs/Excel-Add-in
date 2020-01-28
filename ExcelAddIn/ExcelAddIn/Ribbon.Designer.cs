namespace ExcelAddIn
{
	partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public Ribbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tab1 = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.importBomExample1Btn = this.Factory.CreateRibbonButton();
			this.importBomExample2Btn = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.group1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.Groups.Add(this.group1);
			this.tab1.Label = "Custom Aras Labs Tab";
			this.tab1.Name = "tab1";
			// 
			// group1
			// 
			this.group1.Items.Add(this.importBomExample1Btn);
			this.group1.Items.Add(this.importBomExample2Btn);
			this.group1.Name = "group1";
			// 
			// importBomExample1Btn
			// 
			this.importBomExample1Btn.Label = "Import BOM Example 1";
			this.importBomExample1Btn.Name = "importBomExample1Btn";
			this.importBomExample1Btn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportBomExample1Btn_Click);
			// 
			// importBomExample2Btn
			// 
			this.importBomExample2Btn.Label = "Import BOM Example 2";
			this.importBomExample2Btn.Name = "importBomExample2Btn";
			this.importBomExample2Btn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportBomExample2Btn_Click);
			// 
			// Ribbon1
			// 
			this.Name = "Ribbon1";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton importBomExample1Btn;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton importBomExample2Btn;
	}

	partial class ThisRibbonCollection
	{
		internal Ribbon Ribbon
		{
			get { return this.GetRibbon<Ribbon>(); }
		}
	}
}
