namespace SharpExtensions
{
	public static class XElementToTreeView
	{
   	 	public static void FromXElement(this TreeView T_View, XElement XElem , TreeNode ParentNode = null)
  		{
			// Create temporarty node
			var TempNode = new TreeNode(XElem.Name.LocalName);
			
			// Determine if recursive loop or initial entry
			if(ParentNode == null)
				
				// add root of tree
				T_View.Nodes.Add(TempNode);
			
			// add to parent tree node
			else ParentNode.Nodes.Add(TempNode);
			
			// for each attribute in the element
			foreach(XAttribute XAttr in XElem.Attributes()){
				
				// create a tree node for the attribute
				var AttrNode = new TreeNode(XAttr.Name.LocalName);
				
				// add the value of the attribute as tree node
				AttrNode.Nodes.Add(new TreeNode(XAttr.Value));
				
				// add the node to the tree
				TempNode.Nodes.Add(AttrNode);
			}
			
			// if Element has no child elements
			if(!XElem.HasElements && !XElem.Value.Equals(string.Empty)) {
				
				// add a node for the value if it hase one
				var ValueNode = new TreeNode("value");
				
				// add the value as a node
				ValueNode.Nodes.Add(new TreeNode(XElem.Value));
				
				// add the node to tree
				TempNode.Nodes.Add(ValueNode);
			}
			
			if(XElem.HasElements) {
				// Make element parent node
				var TempParentNode = TempNode;
				
				// recurse
				XElem.Elements().ForEach(XEChild => FromXElement(T_View, XEChild, TempParentNode));
			}
		}
  	}  
}
