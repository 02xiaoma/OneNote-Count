using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.OneNote;
using System.Xml;

namespace OneNote_Count
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // get the onenote xml data
            Microsoft.Office.Interop.OneNote.Application onenoteApp = new Microsoft.Office.Interop.OneNote.Application();
            string notebookXml;
            onenoteApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml, XMLSchema.xs2013);

            XmlDocument dom = new XmlDocument();
            dom.LoadXml(notebookXml);
            treeView.Nodes.Clear();
            treeView.Nodes.Add(new TreeNode(dom.DocumentElement.LocalName + "(" + dom.DocumentElement.ChildNodes.Count + ")"));
            AddNode(dom.DocumentElement, treeView.Nodes[0]);
        }
        private void AddNode(XmlNode inXmlNode, TreeNode inTreeNode)
        {
            // add child xml nodes of inXmlNode to inTreeNode
            if (inXmlNode.HasChildNodes)
            {
                XmlNodeList nodeList = inXmlNode.ChildNodes;
                foreach (XmlElement xNode in nodeList)
                {
                    string str = xNode.GetAttribute("name");
                    if (string.IsNullOrEmpty(str))
                        str = xNode.LocalName;
                    if (xNode.ChildNodes.Count != 0)
                        str = str + "(" + xNode.ChildNodes.Count + ")";
                    inTreeNode.Nodes.Add(new TreeNode(str));
                    AddNode(xNode, inTreeNode.LastNode);
                }
            }
        }
    }
}
