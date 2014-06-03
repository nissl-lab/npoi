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
 * Tony Qu                  2009.3.4        1.2 beta 1
 * 
 * ==============================================================*/

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Collections;

using NPOI.POIFS.FileSystem;
using NPOI.HSSF.Record;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Record.Aggregates;
using NPOI.SS.Util;
using NPOI.DDF;
using NPOI.HSSF.Record.Chart;

namespace NPOI.Tools.POIFSBrowser
{
    internal class DirectoryTreeNode : AbstractTreeNode
    {
        public DirectoryNode DirectoryNode { get; private set; }
        private bool _isExpanded;

        public DirectoryTreeNode(DirectoryNode dn)
            : base(dn)
        {
            
            this.DirectoryNode = dn;

            ChangeImage("Folder");

            this.Nodes.Add(string.Empty); // Dummy node
        }


        internal void OnExpanded()
        {
            if (!_isExpanded)
            {
                this.TreeView.BeginUpdate();
                this.Nodes.Clear();
                this.Nodes.AddRange(GetChildren(this.DirectoryNode,null));
                
                ChangeImage("FolderOpen");

                this.TreeView.EndUpdate();

                _isExpanded = true;
            }
        }

        internal void OnCollapsed()
        {
            ChangeImage("Folder");
        }

        internal static TreeNode[] GetChildren(DirectoryNode node,object innerDoc)
        {
            var children = new List<AbstractTreeNode>();

            var entries = node.Entries;

            while (entries.MoveNext())
            {
                EntryNode entry = entries.Current as EntryNode;
                AbstractTreeNode treeNode;
                if (entry is DirectoryNode)
                {
                    treeNode = new DirectoryTreeNode(entry as DirectoryNode);

                    var o = entry as DirectoryNode;
                }
                else
                {
                    var o = entry as DocumentNode;

                    treeNode = new DocumentTreeNode(entry as DocumentNode);

                    #region handle Excel BIFF records

                    if (treeNode.Text.ToLower() == "workbook")
                    {
                        HandleWorkbook(treeNode, (HSSFWorkbook)innerDoc);
                    }
                    //else if(treeNode.Text.ToLower() == "worddocument")
                    //{
                    //    HandleWord(treeNode, (HWPFDocument)innerDoc);
                    //}

                    #endregion
                }

                children.Add(treeNode);
            }

            children.Sort();

            return children.ToArray();
        }

        //static void HandleWord(TreeNode treeNode, HWPFDocument hwpf)
        //{
        //    TreeNode paragraphsNode = CreateFolderNode(treeNode, "Paragraphs");
        //    List<PAPX> papxs = hwpf.ParagraphTable.GetParagraphs();
        //    foreach (PAPX papx in papxs)
        //    {
        //        TreeNode tn = new TreeNode("PAPX");
        //        paragraphsNode.Nodes.Add(tn);
        //    }
        //    foreach(FieldsDocumentPart part in Enum.GetValues(typeof(FieldsDocumentPart)))
        //    {
        //        LoadDocumentPart(hwpf, treeNode, part.ToString(), part);
        //    }
        //    TreeNode docTextNode = new TreeNode("DocumentText");
        //    docTextNode.Tag = hwpf.GetDocumentText();
        //    treeNode.Nodes.Add(docTextNode);            
        //}
        static TreeNode CreateFolderNode(TreeNode parent, string name)
        {
            TreeNode newnode = new TreeNode(name);
            newnode.ImageKey = "Folder";
            newnode.SelectedImageKey = "Folder";
            parent.Nodes.Add(newnode);
            return newnode;
        }
        //static void LoadDocumentPart(HWPFDocument hwpf,TreeNode parent,string partName,FieldsDocumentPart part)
        //{

        //    TreeNode mainNode = new TreeNode(partName);
        //    mainNode.ImageKey = "Folder";
        //    mainNode.SelectedImageKey = "Folder";
        //    parent.Nodes.Add(mainNode);
        //    Fields fields = hwpf.GetFields();
        //    foreach (Field field in fields.GetFields(part))
        //    {
        //        TreeNode tn = new TreeNode("Field");
        //        mainNode.Nodes.Add(tn);
        //    }
        //}


        static void HandleWorkbook(TreeNode treeNode,HSSFWorkbook hssfworkbook)
        {

                if (hssfworkbook.NumberOfSheets > 0)
                {
                    treeNode.ImageKey = "Folder";
                    treeNode.SelectedImageKey = "Folder";
                    for (int i = 0; i < hssfworkbook.NumberOfSheets; i++)
                    {
                        string sheettext = string.Format("Sheet {0}", i + 1);
                        TreeNode sheetnode =
                            treeNode.Nodes.Add(sheettext, sheettext, "Folder", "Folder");

                        HSSFSheet hssfsheet=((HSSFSheet)hssfworkbook.GetSheetAt(i));
                        EscherAggregate ea = hssfsheet.DrawingEscherAggregate;
                        IEnumerator iterator1 = hssfsheet.Sheet.Records.GetEnumerator();
                        int chartIndex = 1;
                        while (iterator1.MoveNext())
                        {
                            if (iterator1.Current is BOFRecord)
                            {
                                BOFRecord bof = (BOFRecord)iterator1.Current;
                                if (bof.Type == BOFRecordType.Chart)
                                {
                                    string chartTitle = string.Format("Chart {0}" , chartIndex);
                                    TreeNode chartnode = sheetnode.Nodes.Add(chartTitle, chartTitle, "Folder", "Folder");
                                    chartnode.Nodes.Add(new RecordTreeNode(bof));
                                    while (iterator1.MoveNext())
                                    {
                                        if (iterator1.Current is RecordAggregate)
                                        {
                                            RecordAggregate record = (RecordAggregate)iterator1.Current;
                                            chartnode.Nodes.Add(new RecordAggregateTreeNode(record));
                                        }
                                        else /*if(iterator1.Current is Record)*/
                                        {
                                            chartnode.Nodes.Add(new RecordTreeNode((Record)iterator1.Current));
                                        }
                                        if (iterator1.Current is EOFRecord)
                                            break;
                                    }
                                    chartIndex++;
                                }
                            }
                            else if (iterator1.Current is Record)
                            {
                                Record record = (Record)iterator1.Current;
                                sheetnode.Nodes.Add(new RecordTreeNode(record));
                            }
                            else if (iterator1.Current is RecordAggregate)
                            {
                                RecordAggregate record = (RecordAggregate)iterator1.Current;
                                sheetnode.Nodes.Add(new RecordAggregateTreeNode(record));
                            }
                        }
                        //RecordTreeNode rtn = new DirectoryTreeNode();
                        if (ea != null)
                        {
                            foreach (EscherRecord er in ea.EscherRecords)
                            {
                                sheetnode.Nodes.Add(new EscherRecordTreeNode(er));
                            }
                        }
                    }

                }
                else
                {
                    treeNode.ImageKey = "Binary";
                }
                IEnumerator iterator2 = hssfworkbook.Workbook.Records.GetEnumerator();
                while (iterator2.MoveNext())
                {
                    if (iterator2.Current is Record)     //&& !(iterator2.Current is UnknownRecord))
                    {
                        Record record = (Record)iterator2.Current;
                        if (record is DrawingGroupRecord)
                        {
                            hssfworkbook.GetAllPictures();
                        }
                        treeNode.Nodes.Add(new RecordTreeNode(record));
                    }
                    else if (iterator2.Current is RecordBase)
                    {
                        RecordBase record = (RecordBase)iterator2.Current;
                        treeNode.Nodes.Add(record.GetType().Name);
                    }
                }


        }

        private class RecordVisitor1 : RecordVisitor
        {
            private TreeNode node;
            public RecordVisitor1(TreeNode psNode)
            {
                node = psNode;
            }
            #region RecordVisitor 成员

            public void VisitRecord(Record r)
            {
                node.Nodes.Add(new RecordTreeNode(r));
            }

            #endregion
        }

    }
}