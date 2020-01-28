using System;
using System.Xml.Serialization;

namespace ExcelAddIn.Model
{
	[Serializable]
	public class CustomBomStructureEx2 : CustomBomStructure
	{
		public CustomBomStructureEx2() { }

		[XmlIgnore]
		public string Parent { get; set; }
	}
}
