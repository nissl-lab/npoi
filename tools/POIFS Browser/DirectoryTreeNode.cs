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

        internal static TreeNode[] GetChildren(DirectoryNode node,HSSFWorkbook wb)
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
                        HandleWorkbook(treeNode, wb);
                    }

                    #endregion
                }

                children.Add(treeNode);
            }

            children.Sort();

            return children.ToArray();
        }

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

                        IEnumerator iterator1 = ((HSSFSheet)hssfworkbook.GetSheetAt(i)).Sheet.Records.GetEnumerator();
                        while (iterator1.MoveNext())
                        {
                            if (iterator1.Current is Record)
                            {
                                Record record = (Record)iterator1.Current;
                                sheetnode.Nodes.Add(new RecordTreeNode(record));
                            }
                            else if (iterator1.Current is RecordAggregate)
                            {
                                RecordAggregate record = (RecordAggregate)iterator1.Current;

                                TreeNode newnode = new TreeNode(record.GetType().Name);
                                newnode.ImageKey = "Folder";
                                newnode.SelectedImageKey = "Folder";

                                if (record is RowRecordsAggregate)
                                {

                                    IEnumerator recordenum = ((RowRecordsAggregate)record).GetEnumerator();
                                    while (recordenum.MoveNext())
                                    {
                                        if (recordenum.Current is RowRecord)
                                        {
                                            newnode.Nodes.Add(new RecordTreeNode((RowRecord)recordenum.Current));
                                        }
                                    }
                                    CellValueRecordInterface[] valrecs = ((RowRecordsAggregate)record).GetValueRecords();
                                    for (int j = 0; j < valrecs.Length; j++)
                                    {
                                        CellValueRecordTreeNode cvrtn=new CellValueRecordTreeNode(valrecs[j]);
                                         if (valrecs[j] is FormulaRecordAggregate)
                                        {
                                            FormulaRecordAggregate fra = ((FormulaRecordAggregate)valrecs[j]);
                                            cvrtn.ImageKey = "Folder";
                                             if(fra.FormulaRecord!=null)
                                                cvrtn.Nodes.Add(new RecordTreeNode(fra.FormulaRecord));
                                             if(fra.StringRecord!=null)
                                                cvrtn.Nodes.Add(new RecordTreeNode(fra.StringRecord));
                                        }
                                        newnode.Nodes.Add(cvrtn);
                                    }
                                }
                                else if (record is ColumnInfoRecordsAggregate)
                                {
                                    IEnumerator recordenum = ((ColumnInfoRecordsAggregate)record).GetEnumerator();
                                    while (recordenum.MoveNext())
                                    {
                                        if (recordenum.Current is ColumnInfoRecord)
                                        {
                                            newnode.Nodes.Add(new RecordTreeNode((ColumnInfoRecord)recordenum.Current));
                                        }
                                    }
                                }
                                else if (record is PageSettingsBlock)
                                {
                                    IEnumerator recordenum = ((PageSettingsBlock)record).GetEnumerator();
                                    while (recordenum.MoveNext())
                                    {
                                        if (recordenum.Current is Record)
                                        {
                                            newnode.Nodes.Add(new RecordTreeNode((Record)recordenum.Current));
                                        }
                                    }
                                }
                                else if (record is MergedCellsTable)
                                {
                                    foreach(CellRangeAddress subRecord in ((MergedCellsTable)record).MergedRegions)
                                    {
                                        newnode.Nodes.Add(new CellRangeAddressTreeNode(subRecord));
                                    }
                                }
                                else if (record is ConditionalFormattingTable)
                                {
                                    ConditionalFormattingTable cft = (ConditionalFormattingTable)record;
                                    for (int j = 0; j < cft.Count; j++)
                                    {
                                        CFRecordsAggregate cfra = cft.Get(i);

                                        TreeNode headernode = new RecordTreeNode(cfra.Header);
                                        for (int k = 0; k < cfra.NumberOfRules; k++)
                                        {
                                            newnode.Nodes.Add(new RecordTreeNode(cfra.GetRule(k)));
                                        }
                                        newnode.Nodes.Add(headernode);
                                    }
                                }
                                else
                                {
                                    newnode = new TreeNode(record.GetType().Name);
                                }
                                sheetnode.Nodes.Add(newnode);
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

    }
}