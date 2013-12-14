using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
// add name space
using System.Xml;
using System.IO;

namespace NPOI.StyleViewer
{
    public partial class Default : Form
    {
        public Default()
        {
            InitializeComponent();
        }

        private void Default_Load(object sender, EventArgs e)
        {
            if (StyleDocumentInitializer.Load() && StyleDocumentInitializer.IsReady)
            {
                foreach (WordStyle style in StyleDocumentInitializer.StyleSet.OrderBy(p => p.TypeName).ToList())
                {
                    TreeNode T = new TreeNode(style.TypeName);

                    foreach (var item in style.Styles.OrderBy(p=>p.Key).ToList())
                    {
                        T.Nodes.Add(item.Key);
                    }

                    tvWordStyle.Nodes.Add(T);
                }
            }

            tvWordStyle.ExpandAll();
        }

        private void tvWordStyle_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (tvWordStyle.SelectedNode.Parent != null)
            {
                var TypeName = tvWordStyle.SelectedNode.Parent.Text;
                var StyleID = tvWordStyle.SelectedNode.Text;

                rtxtStyleXml.Text = StyleDocumentInitializer.GetXmlByStyleID(TypeName, StyleID);

            }
            else
            {
                rtxtStyleXml.Text = string.Empty;
            }
        }
    }
}
