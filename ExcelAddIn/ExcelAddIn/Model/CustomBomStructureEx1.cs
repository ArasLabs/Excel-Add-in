using System;
using System.Xml.Serialization;

namespace ExcelAddIn.Model
{
	[Serializable]
	public class CustomBomStructureEx1 : CustomBomStructure
	{
		public CustomBomStructureEx1() { }

		[XmlIgnore]
		public int Level { get; set; }
	}
}
