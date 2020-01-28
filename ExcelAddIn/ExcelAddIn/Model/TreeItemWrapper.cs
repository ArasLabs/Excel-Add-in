using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace ExcelAddIn.Model
{
	[Serializable]
	public class TreeItemWrapper<T>
	{
		public TreeItemWrapper() { }
		public TreeItemWrapper(List<TreeItem<T>> leafNodes)
		{
			LeafNodes = leafNodes;
		}

		[XmlElement("LeafNodesOfCustomBomStructure")]
		public List<TreeItem<T>> LeafNodes { get; }
	}
}
