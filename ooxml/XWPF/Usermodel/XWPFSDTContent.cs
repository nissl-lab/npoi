/* ====================================================================
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
                if (o is CT_P)
                {
                    XWPFParagraph p = new XWPFParagraph((CT_P)o, part);
                    bodyElements.Add(p);
                    paragraphs.Add(p);
                }
                else if (o is CT_Tbl)
                {
                    XWPFTable t = new XWPFTable((CT_Tbl)o, part);
                    bodyElements.Add(t);
                    tables.Add(t);
                }
                else if (o is CT_SdtBlock)
                {
                    XWPFSDT c = new XWPFSDT(((CT_SdtBlock)o), part);
                    bodyElements.Add(c);
                    contentControls.Add(c);
                }
                else if (o is CT_R)
                {
                    XWPFRun run = new XWPFRun((CT_R)o, parent);
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
                    if (o is XWPFParagraph)
                    {
                        AppendParagraph((XWPFParagraph)o, text);
                        addNewLine = true;
                    }
                    else if (o is XWPFTable)
                    {
                        AppendTable((XWPFTable)o, text);
                        addNewLine = true;
                    }
                    else if (o is XWPFSDT)
                    {
                        text.Append(((XWPFSDT)o).Content.Text);
                        addNewLine = true;
                    }
                    else if (o is XWPFRun)
                    {
                        text.Append(((XWPFRun)o).ToString());
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
                    if (cell is XWPFTableCell)
                    {
                        text.Append(((XWPFTableCell)cell).GetTextRecursively());
                    }
                    else if (cell is XWPFSDTCell)
                    {
                        text.Append(((XWPFSDTCell)cell).Content.Text);
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