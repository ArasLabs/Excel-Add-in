using ExcelAddIn.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAddIn.Helpers
{
	public static class GenericTreeGenerator
	{
		public static List<TreeItem<T>> GenerateTreeFromList<T, K>(this List<T> leafNodes,
			Func<T, K> nodeIdSelector,
			Func<T, K> parentNodeIdSelector,
			Func<T, K, bool> breakCondition = default(Func<T, K, bool>),
			K rootNodeId = default(K),
			int startIndex = default(int))
		{
			List<T> specialLeafNodes = leafNodes
				.Where((val, index) => index >= startIndex && parentNodeIdSelector(val).Equals(rootNodeId))
				.ToList();

			List<TreeItem<T>> result = new List<TreeItem<T>>(specialLeafNodes.Count());

			foreach (var leafNode in specialLeafNodes)
			{
				if (breakCondition?.Invoke(leafNode, parentNodeIdSelector(leafNode)) == true)
					break;

				result.Add(new TreeItem<T>
				{
					ParentNode = leafNode,
					LeafNodes = leafNodes.GenerateTreeFromList(
						nodeIdSelector,
						parentNodeIdSelector,
						breakCondition,
						nodeIdSelector(leafNode),
						leafNodes.IndexOf(leafNode)
					).ToList()
				});
			}

			return result;
		}
	}
}
