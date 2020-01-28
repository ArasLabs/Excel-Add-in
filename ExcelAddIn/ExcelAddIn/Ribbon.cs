using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using ExcelAddIn.Helpers;
using ExcelAddIn.Model;
using System.Xml.Serialization;
using System.Windows.Forms;

namespace ExcelAddIn
{
	public partial class Ribbon
	{
		private void Ribbon_Load(object sender, RibbonUIEventArgs e)
		{

		}

		#region Example 1
		private void ImportBomExample1Btn_Click(object sender, RibbonControlEventArgs e)
		{
			try
			{
				var treeBomStructure = CreateDataModelForStructureFromExample1()
					.GenerateTreeFromList(x => x.Level + 1, x => x.Level, (x, y) => x.Level < y);

				var body = CustomXmlSerializer.SerializeObject(new TreeItemWrapper<CustomBomStructureEx1>(treeBomStructure), new XmlRootAttribute("ArrayOfLeafNodesOfCustomBomStructure"));
				var res = Globals.ThisAddIn.Innovator.applyMethod("TestExcelAddInExample", body);
				if (res.isError())
					throw new Exception(res.getErrorString());

				MessageBox.Show("Bom Structure was successfully created!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.Message, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private List<CustomBomStructureEx1> CreateDataModelForStructureFromExample1()
		{
			try
			{
				Worksheet currentWorkSheet = Globals.ThisAddIn.GetActiveWorkSheet();
				var activeApp = Globals.ThisAddIn.GetActiveApplication();

				if (currentWorkSheet.UsedRange.Columns.Count != 4
					|| (string)(currentWorkSheet.Cells[1, 1] as Range).Value != "Level"
					|| (string)(currentWorkSheet.Cells[1, 2] as Range).Value != "P/N"
					|| (string)(currentWorkSheet.Cells[1, 3] as Range).Value != "Name"
					|| (string)(currentWorkSheet.Cells[1, 4] as Range).Value != "Quantity")
				{
					throw new Exception("Data structure is incorrect, please normalize your structure to ( Level(int), P/N(string), Name(string), Quantity(string) )");
				}

				int sRowCount = currentWorkSheet.UsedRange.Rows.Count;
				List<CustomBomStructureEx1> bomStructure = new List<CustomBomStructureEx1>(sRowCount);
				Range level = currentWorkSheet.Range["A2:A" + sRowCount];
				Range partNumber = currentWorkSheet.Range["B2:B" + sRowCount];
				Range name = currentWorkSheet.Range["C2:C" + sRowCount];
				Range quantity = currentWorkSheet.Range["D2:D" + sRowCount];

				// Convert excel data to data model
				for (int i = 0; i < sRowCount - 1; ++i)
				{
					bomStructure.Add(
						new CustomBomStructureEx1
						{
							Level = Convert.ToInt32(GetColumnValues(level)[i]),
							Quantity = GetColumnValues(quantity)[i],
							PartNumber = GetColumnValues(partNumber)[i],
							Name = GetColumnValues(name)[i]
						});
				}

				return bomStructure;
			}
			catch
			{
				throw new Exception("Error by converting data structure for first example ( Level(int), P/N(string), Name(string), Quantity(string) )");
			}
		}
		#endregion

		#region Example 2
		private void ImportBomExample2Btn_Click(object sender, RibbonControlEventArgs e)
		{
			try
			{
				var treeBomStructure = CreateDataModelForStructureFromExample2()
					.GenerateTreeFromList(x => x.PartNumber, x => x.Parent, rootNodeId: "-");

				var body = CustomXmlSerializer.SerializeObject(new TreeItemWrapper<CustomBomStructureEx2>(treeBomStructure), new XmlRootAttribute("ArrayOfLeafNodesOfCustomBomStructure"));
				var res = Globals.ThisAddIn.Innovator.applyMethod("TestExcelAddInExample", body);
				if (res.isError())
					throw new Exception(res.getErrorString());

				MessageBox.Show("Bom Structure was successfully created!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.Message, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private List<CustomBomStructureEx2> CreateDataModelForStructureFromExample2()
		{
			try
			{
				Worksheet currentWorkSheet = Globals.ThisAddIn.GetActiveWorkSheet();
				var activeApp = Globals.ThisAddIn.GetActiveApplication();

				int sRowCount = currentWorkSheet.UsedRange.Rows.Count;

				if (currentWorkSheet.UsedRange.Columns.Count != 4
					|| (string)(currentWorkSheet.Cells[1, 1] as Range).Value != "Parent"
					|| (string)(currentWorkSheet.Cells[1, 2] as Range).Value != "Child"
					|| (string)(currentWorkSheet.Cells[1, 3] as Range).Value != "Qty"
					|| (string)(currentWorkSheet.Cells[1, 4] as Range).Value != "Name")
				{
					throw new Exception("Data structure is incorrect, please normalize your structure to ( Parent(string), Child(string), Qty(string), Name(string) )");
				}

				List<CustomBomStructureEx2> bomStructure = new List<CustomBomStructureEx2>(sRowCount);
				Range parent = currentWorkSheet.Range["A2:A" + sRowCount];
				Range partNumber = currentWorkSheet.Range["B2:B" + sRowCount];
				Range quantity = currentWorkSheet.Range["C2:C" + sRowCount];
				Range name = currentWorkSheet.Range["D2:D" + sRowCount];

				// Convert excel data to data model
				for (int i = 0; i < sRowCount - 1; ++i)
				{
					bomStructure.Add(
						new CustomBomStructureEx2
						{
							Quantity = GetColumnValues(quantity)[i],
							PartNumber = GetColumnValues(partNumber)[i],
							Parent = GetColumnValues(parent)[i],
							Name = GetColumnValues(name)[i]
						});
				}

				return bomStructure;
			}
			catch
			{
				throw new Exception("Error by converting data structure for first example ( Level(int), P/N(string), Name(string), Quantity(string) )");
			}
		}
		#endregion

		#region Helpers
		private List<string> GetColumnValues(Range level)
		{
			object[,] cellValues = (object[,])level.Value2;
			return cellValues.Cast<object>()
							 .Select(o => o.ToString())
							 .ToList();
		}
		#endregion
	}
}
