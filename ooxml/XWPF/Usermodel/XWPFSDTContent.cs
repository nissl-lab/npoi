﻿/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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
using System;
using NPOI.OpenXmlFormats.Wordprocessing;
using System.Text; 
using Cysharp.Text;
using System.Collections.Generic;
namespace NPOI.XWPF.UserModel
{
    /**
     * Experimental class to offer rudimentary Read-only Processing of 
     *  of the contentblock of an SDT/ContentControl.
     *  
     *
     *
     * WARNING - APIs expected to change rapidly
     * 
     */
    public class XWPFSDTContent : ISDTContent
    {

        // private IBody part;
        // private XWPFDocument document;
        private List<XWPFParagraph> paragraphs = new List<XWPFParagraph>();
        private List<XWPFTable> tables = new List<XWPFTable>();
        private List<XWPFRun> runs = new List<XWPFRun>();
        private List<XWPFSDT> contentControls = new List<XWPFSDT>();
        private List<ISDTContents> bodyElements = new List<ISDTContents>();

        public XWPFSDTContent(CT_SdtContentRun sdtRun, IBody part, IRunBody parent)
        {
            foreach (CT_R ctr in sdtRun.GetRList())
            {
                XWPFRun run = new XWPFRun((CT_R)ctr, parent);
                runs.Add(run);
                bodyElements.Add(run);
            }
        }
        public XWPFSDTContent(CT_SdtContentBlock block, IBody part, IRunBody parent)
        {
            
            foreach (object o in block.Items)
            {
                if (o is CT_P ctP)
                {
                    XWPFParagraph p = new XWPFParagraph(ctP, part);
                    bodyElements.Add(p);
                    paragraphs.Add(p);
                }
                else if (o is CT_Tbl tbl)
                {
                    XWPFTable t = new XWPFTable(tbl, part);
                    bodyElements.Add(t);
                    tables.Add(t);
                }
                else if (o is CT_SdtBlock sdtBlock)
                {
                    XWPFSDT c = new XWPFSDT(sdtBlock, part);
                    bodyElements.Add(c);
                    contentControls.Add(c);
                }
                else if (o is CT_R r)
                {
                    XWPFRun run = new XWPFRun(r, parent);
                    runs.Add(run);
                    bodyElements.Add(run);
                }
            }
        }

        public String Text
        {
            get
            {
                StringBuilder text = new StringBuilder();
                bool addNewLine = false;
                for (int i = 0; i < bodyElements.Count; i++)
                {
                    Object o = bodyElements[i];
                    if (o is XWPFParagraph paragraph)
                    {
                        AppendParagraph(paragraph, text);
                        addNewLine = true;
                    }
                    else if (o is XWPFTable table)
                    {
                        AppendTable(table, text);
                        addNewLine = true;
                    }
                    else if (o is XWPFSDT xwpfsdt)
                    {
                        text.Append(xwpfsdt.Content.Text);
                        addNewLine = true;
                    }
                    else if (o is XWPFRun run)
                    {
                        text.Append(run.ToString());
                        addNewLine = false;
                    }
                    if (addNewLine && i < bodyElements.Count-1)
                    {
                        text.Append("\n");
                    }
                }
                return text.ToString();
            }
        }

        private void AppendTable(XWPFTable table, StringBuilder text)
        {
            //this works recursively to pull embedded tables from within cells
            foreach (XWPFTableRow row in table.Rows)
            {
                List<ICell> cells = row.GetTableICells();
                for (int i = 0; i < cells.Count; i++)
                {
                    ICell cell = cells[i];
                    if (cell is XWPFTableCell tableCell)
                    {
                        text.Append(tableCell.GetTextRecursively());
                    }
                    else if (cell is XWPFSDTCell xwpfsdtCell)
                    {
                        text.Append(xwpfsdtCell.Content.Text);
                    }
                    if (i < cells.Count - 1)
                    {
                        text.Append("\t");
                    }
                }
                text.Append('\n');
            }
        }

        private void AppendParagraph(XWPFParagraph paragraph, StringBuilder text)
        {
            foreach (IRunElement run in paragraph.Runs)
            {
                text.Append(run.ToString());
            }
        }


        public override String ToString()
        {
            return this.Text;
        }
    }

}