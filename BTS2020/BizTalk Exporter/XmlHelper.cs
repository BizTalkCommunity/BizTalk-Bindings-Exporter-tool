using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml;

namespace BizTalk_Exporter
{
    public class XmlHelper
    {
        internal TreeNode FindParent(TreeView treeView, string text)
        {   //Help the child find his parents.
            foreach (TreeNode node in treeView.Nodes)
            {
                if (node.Text == text)
                    return node;
            }
            return null;
        }
        internal void AddRootAndChild(TreeView selectedTree, TreeView destinyTree)
        {
            var node = destinyTree.Nodes.Add(selectedTree.SelectedNode.Text);
            node.Name = selectedTree.SelectedNode.Name;
            foreach (TreeNode subnodeParent in selectedTree.SelectedNode.Nodes)
            {
                var subnode = node.Nodes.Add(subnodeParent.Text);
                subnode.Name = subnodeParent.Name;
            }
            node.Expand(); //we want to see what's inside
        }
        internal bool FindSubNodeInNode(TreeNode parent, string text)
        {
            foreach (TreeNode subnode in parent.Nodes)
            {
                if (subnode.Text == text)
                    return true;
            }
            return false;
        }
        internal void RemoveExcessBindings(string path, bool receive, List<string> portsList)
        {   //it's like lipossuction, but for a fat XML
            try
            {
                XmlDocument bindings = new XmlDocument();
                bindings.Load(path);
                bindings.SelectSingleNode("//BindingInfo//ModuleRefCollection").RemoveAll();

                if (receive)
                {   //a little trim here
                    bindings.SelectSingleNode("//BindingInfo//SendPortCollection").RemoveAll();
                    //loop ReceivePorts and remove unselected ones
                    foreach (XmlNode receivePort in bindings.SelectNodes("//BindingInfo//ReceivePortCollection//ReceivePort"))
                    {
                        if (!portsList.Contains(receivePort.Attributes["Name"].Value))
                            receivePort.ParentNode.RemoveChild(receivePort);
                    }
                }
                else
                {   //a little trim there
                    bindings.SelectSingleNode("//BindingInfo//ReceivePortCollection").RemoveAll();
                    //loop SendPorts and remove unselected ones
                    foreach (XmlNode sendPort in bindings.SelectNodes("//BindingInfo//SendPortCollection//SendPort"))
                    {
                        if (!portsList.Contains(sendPort.Attributes["Name"].Value))
                            sendPort.ParentNode.RemoveChild(sendPort);
                    }
                }
                    bindings.Save(path);
            }
            catch (Exception ex)
            { throw ex; }
        }
        internal void RemoveExcessBindings(string originalPath, string path, bool receive, List<string> portsList)
        {   //it's like lipossuction, but for a fat XML
            try
            {
                XmlDocument bindings = new XmlDocument();
                bindings.Load(originalPath);
                bindings.SelectSingleNode("//BindingInfo//ModuleRefCollection").RemoveAll();

                if (receive)
                {   //a little trim here
                    bindings.SelectSingleNode("//BindingInfo//SendPortCollection").RemoveAll();
                    //loop ReceivePorts and remove unselected ones
                    foreach (XmlNode receivePort in bindings.SelectNodes("//BindingInfo//ReceivePortCollection//ReceivePort"))
                    {
                        if (!portsList.Contains(receivePort.Attributes["Name"].Value))
                            receivePort.ParentNode.RemoveChild(receivePort);
                    }
                }
                else
                {   //a little trim there
                    bindings.SelectSingleNode("//BindingInfo//ReceivePortCollection").RemoveAll();
                    //loop SendPorts and remove unselected ones
                    foreach (XmlNode sendPort in bindings.SelectNodes("//BindingInfo//SendPortCollection//SendPort"))
                    {
                        if (!portsList.Contains(sendPort.Attributes["Name"].Value))
                            sendPort.ParentNode.RemoveChild(sendPort);
                    }
                }
                bindings.Save(path);
            }
            catch (Exception ex)
            { throw ex; }
        }
        internal void ReplaceEnvironmentBindings(List<excelData> ports, string path, string type)
        {
            XmlDocument bindings = new XmlDocument();
            bindings.Load(path);
            foreach (excelData port in ports)
            {
                XmlNode node = SearchNode(port.portName, bindings.SelectSingleNode("//BindingInfo").ChildNodes);
                if (node != null)
                {   //dig deeper
                    if (type == "Receive")
                        node.SelectSingleNode("//ReceiveLocations//Address").InnerText = port.portURI;
                    else
                        node.SelectSingleNode("//PrimaryTransport//Address").InnerText = port.portURI;
                }
            }
            bindings.Save(path);
        }
        internal XmlNode SearchNode(string portName, XmlNodeList list)
        {
            XmlNode returnNode;
            foreach (XmlNode node in list)
            {
                if (node.HasChildNodes)
                {
                    returnNode = SearchSubNode(portName, node.ChildNodes);
                    if (returnNode != null)
                        return returnNode;
                }
            }
            return null;
        }
        internal XmlNode SearchSubNode(string portName, XmlNodeList nodeList)
        {
            foreach (XmlNode node in nodeList)
            {
                if (node.Attributes != null)
                    if (node.Attributes["Name"] != null)
                    {
                        if (node.Attributes["Name"].Value == portName)
                            return node;
                    }
            }
            return null;
        }
        internal void CollectChildren(TreeNode node, ref List<string> children)
        {   //should have named it better... toss them all into a bag
            if (node.Nodes.Count == 0)
                children.Add(node.Text);
            else
                foreach (TreeNode n in node.Nodes)
                    CollectChildren(n, ref children);
        }
    }
}
