using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace ExcelAddIn.Model
{
	[Serializable]
	public class TreeItem<T>
	{
		public TreeItem() { }


		[XmlElement(ElementName = "ParendNodeOfCustomBomStructure")]
		public T ParentNode { get; set; }
		
		[XmlElement(ElementName = "LeafNodesOfCustomBomStructure")]
		public List<TreeItem<T>> LeafNodes { get; set; }
	}
}
