using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Model
{
	[Serializable]
	public class CustomBomStructure
	{
		public CustomBomStructure() { }
		
		public string PartNumber { get; set; }
		public string Name { get; set; }
		public string Quantity { get; set; }
	}
}
