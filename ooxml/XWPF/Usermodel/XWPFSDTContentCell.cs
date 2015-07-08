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
namespace NPOI.XWPF.UserModel
{
    using System;
    using System.Collections;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Text;
    using System.IO;
    using System.Xml;

    /**
     * Experimental class to offer rudimentary Read-only Processing of
     * of the XWPFSDTCellContent.
     * <p/>
     * WARNING - APIs expected to change rapidly
     */
    public class XWPFSDTContentCell : ISDTContent
    {

        //A full implementation would grab the icells
        //that a content cell can Contain.  This would require
        //significant Changes, including changing the notion that the
        //parent of a cell can be not just a row, but an sdt.
        //For now we are just grabbing the text out of the text tokentypes.

        //private List<ICell> cells = new List<ICell>().

        private String text = "";

        public XWPFSDTContentCell(CT_SdtContentCell sdtContentCell,
                                  XWPFTableRow xwpfTableRow, IBody part)
        {
            StringBuilder sb = new StringBuilder();
            

            //keep track of the following,
            //and add "\n" only before the start of a body
            //element if it is not the first body element.

            //index of cell in row
            int tcCnt = 0;
            //count of body objects
            int iBodyCnt = 0;
            int depth = 1;
            string sdtXml = sdtContentCell.ToString();
            using (StringReader sr = new StringReader(sdtXml))
            {
                XmlParserContext context = new XmlParserContext(null, POIXMLDocumentPart.NamespaceManager,
                    null, XmlSpace.Preserve);
                using (XmlReader cursor = XmlReader.Create(sr, null, context))
                {
                    while (cursor.Read() && depth > 0)
                    {
                        if (cursor.NodeType == XmlNodeType.Text)
                        {
                            sb.Append(cursor.ReadContentAsString());
                        }
                        else if (IsStartToken(cursor, "tr"))
                        {
                            tcCnt = 0;
                            iBodyCnt = 0;
                        }
                        else if (IsStartToken(cursor, "tc"))
                        {
                            if (tcCnt++ > 0)
                            {
                                sb.Append("\t");
                            }
                            iBodyCnt = 0;
                        }
                        else if (IsStartToken(cursor, "p") ||
                              IsStartToken(cursor, "tbl") ||
                              IsStartToken(cursor, "sdt"))
                        {
                            if (iBodyCnt > 0)
                            {
                                sb.Append("\n");
                            }
                            iBodyCnt++;
                        }

                    }
                }
            }
            //IEnumerator cursor = sdtContentCell.Items.GetEnumerator();
            //while (cursor.MoveNext() && depth > 0)
            //{
            //    //TokenType t = cursor.ToNextToken();
            //    object t = cursor.Current;
            //    if (t is CT_Text)//??
            //    {
            //        //sb.Append(cursor.TextValue);
            //    }
            //    else if (IsStartToken(cursor, "tr"))
            //    {
            //        tcCnt = 0;
            //        iBodyCnt = 0;
            //    }
            //    else if (IsStartToken(cursor, "tc"))
            //    {
            //        if (tcCnt++ > 0)
            //        {
            //            sb.Append("\t");
            //        }
            //        iBodyCnt = 0;
            //    }
            //    else if (IsStartToken(cursor, "p") ||
            //          IsStartToken(cursor, "tbl") ||
            //          IsStartToken(cursor, "sdt"))
            //    {
            //        if (iBodyCnt > 0)
            //        {
            //            sb.Append("\n");
            //        }
            //        iBodyCnt++;
            //    }
            //    //if (cursor.IsStart())
            //    //{
            //    //    depth++;
            //    //}
            //    //else if (cursor.IsEnd())
            //    //{
            //    //    depth--;
            //    //}
            //}
            text = sb.ToString();
        }
        private bool IsStartToken(XmlReader cursor, String string1)
        {
            if (!cursor.IsStartElement())
            {
                return false;
            }

            if (cursor.LocalName == string1)
            {
                return true;
            }
            return false;
        }

        private bool IsStartToken(object cursor, String string1)
        {
            throw new NotImplementedException();
            //if (!cursor.IsStart())
            //{
            //    return false;
            //}
            //QName qName = cursor.Name;
            //if (qName != null && qName.LocalPart != null &&
            //        qName.LocalPart.Equals(string1))
            //{
            //    return true;
            //}
            return false;
        }


        public string Text
        {
            get
            {
                return text;
            }
        }

        public override string ToString()
        {
            return Text;
        }
    }
}

