/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * POIFS Browser 
 * Author: NPOI Team
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * Huseyin Tufekcilerli     2008.11         1.0
 * Tony Qu                  2009.2.18       1.2 alpha
 * 
 * ==============================================================*/

using System;
using System.IO;
using System.Collections;
using System.Windows.Forms;

using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Record;

namespace NPOI.Tools.POIFSBrowser
{
    public partial class MainForm : Form
    {
        private POIFSFileSystem _currentFileSystem;


        public MainForm()
        {
            InitializeComponent();
        }

        private void OpenDocument(string path)
        {
            HSSFWorkbook hssfworkbook = null;
            //HWPFDocument hwpf = null;

            using (var stream = File.OpenRead(path))
            {
                try
                {
                    _currentFileSystem = new POIFSFileSystem(stream);
                    //supposing every Excel file has .xls as the extension 
                    if (path.ToLower().IndexOf(".xls")>0)
                    {
                        hssfworkbook = new HSSFWorkbook(_currentFileSystem);
                    }
                    //else if (path.ToLower().IndexOf(".doc") > 0)
                    //{
                    //    hwpf =new HWPFDocument(_currentFileSystem);
                    //}
                }
                catch (Exception)
                {
                    MessageBox.Show("Error opening file. Possibly the file is not an OLE2 Compund file.",
                        "Open File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if (_currentFileSystem != null)
            {
                this.Text = string.Format("POIFS Browser - {0}", path);
                documentTreeView.BeginUpdate();
                documentTreeView.Nodes.Clear();

                TreeNode[] children=null;
                if (hssfworkbook != null)
                {
                    children = DirectoryTreeNode.GetChildren(_currentFileSystem.Root, hssfworkbook);
                }
                else
                {
                    children = DirectoryTreeNode.GetChildren(_currentFileSystem.Root, null);
                }
                documentTreeView.Nodes.AddRange(children);
                documentTreeView.EndUpdate();

            }
        }

        private void documentTreeView_AfterExpand(object sender, TreeViewEventArgs e)
        {
            var directoryTreeNode = e.Node as DirectoryTreeNode;

            if (directoryTreeNode != null)
            {
                directoryTreeNode.OnExpanded();
            }
        }

        private void documentTreeView_AfterCollapse(object sender, TreeViewEventArgs e)
        {
            var directoryTreeNode = e.Node as DirectoryTreeNode;

            if (directoryTreeNode != null)
            {
                directoryTreeNode.OnCollapsed();
            }
        }

        private void documentTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            ClearViewControls();

            if (e.Node is DocumentTreeNode)
            {
                var documentTreeNode = e.Node as DocumentTreeNode;
                var stream = documentTreeNode.GetDocumentStream();

                var viewableArray = documentTreeNode.DocumentNode.Document.ViewableArray;
                if (viewableArray.Length > 0)
                {
                    viewableTextBox.Text = viewableArray.GetValue(0).ToString();
                }

                if (PropertySet.IsPropertySetStream(stream))
                {
                    var ps = PropertySetFactory.Create(stream);

                    propertiesListView.Items.AddRange(PropertyListViewItem.Create(ps));

                    if (!streamTabControl.TabPages.ContainsKey("tabPageProperties"))
                    {
                        streamTabControl.TabPages.Add(tabPageProperties);
                    }
                }
                else
                {
                    if (streamTabControl.TabPages.ContainsKey("tabPageProperties"))
                    {
                        streamTabControl.TabPages.Remove(tabPageProperties);
                    }
                }

                if (!streamTabControl.TabPages.ContainsKey("tabPageBinary"))
                {
                    streamTabControl.TabPages.Add(tabPageBinary);
                }

            }
            else if (e.Node is AbstractRecordTreeNode)
            {
                AbstractRecordTreeNode node = (AbstractRecordTreeNode)e.Node;

                if (node.HasBinary)
                {
                    byte[] buffer = node.GetBytes();
                    viewableTextBox.Text = NPOI.Util.HexDump.Dump(buffer, buffer.Length, 0);
                    if (!streamTabControl.TabPages.ContainsKey("tabPageBinary"))
                    {
                        streamTabControl.TabPages.Add(tabPageBinary);
                    }
                }
                else
                {
                    if (streamTabControl.TabPages.ContainsKey("tabPageBinary"))
                    {
                        streamTabControl.TabPages.Remove(tabPageBinary);
                    }
                }
                if (!streamTabControl.TabPages.ContainsKey("tabPageProperties"))
                {
                    streamTabControl.TabPages.Add(tabPageProperties);
                }    
                propertiesListView.Items.AddRange(node.GetPropertyList());

                if (!streamTabControl.TabPages.ContainsKey("tabPageBinary"))
                {
                    streamTabControl.TabPages.Add(tabPageBinary);
                }
            }
            else
            {
                if (!streamTabControl.TabPages.ContainsKey("tabPageProperties"))
                {
                    streamTabControl.TabPages.Add(tabPageProperties);
                }            
            }

        }

        private void ClearViewControls()
        {
            viewableTextBox.Clear();
            propertiesListView.Items.Clear();
        }

        private string filename;
        private void Open_Click(object sender, EventArgs e)
        {
            var dialog = new OpenFileDialog()
            {
                Filter = "Microsoft Office 97-2003 Documents|*.xls;*.doc;*.ppt|POIFS File|*.poifs|All files (*.*)|*.*",
                Multiselect = false,
                Title = "Open OLE2 Compund Document",
            };

            var result = dialog.ShowDialog(this);

            if (result == DialogResult.OK)
            {
                filename = dialog.FileName;
                OpenDocument(filename);

                ClearViewControls();
            }
        }

        private void saveAsMenu_Click(object sender, EventArgs e)
        {

            var documentTreeNode = documentTreeView.SelectedNode as DocumentTreeNode;
            var directoryTreeNode = documentTreeView.SelectedNode as DirectoryTreeNode;
            if (documentTreeNode != null)
            {
                var dialog = new SaveFileDialog()
                {
                    Filter = "All files|*.*",
                    Title = "Save Document Stream As"
                };

                var result = dialog.ShowDialog(this);

                if (result == DialogResult.OK)
                {
                    using (var stream = documentTreeNode.GetDocumentStream())
                    using (var fileStream = File.OpenWrite(dialog.FileName))
                    {
                        var bufferLength = 4096;
                        var buffer = new byte[bufferLength];
                        int bytesRead;

                        while ((bytesRead = stream.Read(buffer, 0, bufferLength)) > 0)
                        {
                            fileStream.Write(buffer, 0, bytesRead);
                        }
                    } 
                }
            }
        }

        private void aboutMenuItem_Click(object sender, EventArgs e)
        {
            AboutDialog.Instance.ShowDialog(this);
        }

        private void fileExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void refreshToolScripButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filename))
                return; 

            OpenDocument(filename);
            ClearViewControls();
        }
    }
}
