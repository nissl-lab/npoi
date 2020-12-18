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

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using NPOI.HWPF.Model;
using NPOI.HWPF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.Util;


namespace NPOI.HWPF.Converter
{
    public class AbstractWordUtils
    {
        static POILogger logger = POILogFactory.GetLogger(typeof(AbstractWordUtils));
        public static float TWIPS_PER_INCH = 1440.0f;
        public static int TWIPS_PER_PT = 20;


        /**
     * Creates array of all possible cell edges. In HTML (and FO) cells from
     * different rows and same column should have same width, otherwise spanning
     * shall be used.
     * 
     * @param table
     *            table to build cell edges array from
     * @return array of cell edges (including leftest one) in twips
     */
        public static int[] BuildTableCellEdgesArray(Table table)
        {
            SortedList<int, int> edges = new SortedList<int, int>();

            for (int r = 0; r < table.NumRows; r++)
            {
                TableRow tableRow = table.GetRow(r);
                for (int c = 0; c < tableRow.NumCells(); c++)
                {
                    TableCell tableCell = tableRow.GetCell(c);
                    if (!edges.ContainsKey(tableCell.GetLeftEdge()))
                        edges.Add(tableCell.GetLeftEdge(), 0);
                    if (!edges.ContainsKey(tableCell.GetLeftEdge() + tableCell.GetWidth()))
                        edges.Add(tableCell.GetLeftEdge() + tableCell.GetWidth(), 0);
                }
            }

            int[] result = new int[edges.Count];

            edges.Keys.CopyTo(result, 0);
            return result;
        }
        static bool CanBeMerged(XmlNode node1, XmlNode node2, String requiredTagName)
        {
            if (node1.NodeType != XmlNodeType.Element || node2.NodeType != XmlNodeType.Element)
                return false;

            XmlElement element1 = (XmlElement)node1;
            XmlElement element2 = (XmlElement)node2;

            if (!StringEquals(requiredTagName, element1.Name)
                    || !StringEquals(requiredTagName, element2.Name))
                return false;

            if (element1.Attributes.Count != element2.Attributes.Count)
                return false;

            for (int i = 0; i < element1.Attributes.Count; i++)
            {
                XmlAttribute attr1 = (XmlAttribute)element1.Attributes[i];
                XmlAttribute attr2;
                if (string.IsNullOrEmpty(attr1.NamespaceURI))
                    attr2 = (XmlAttribute)element2.Attributes.GetNamedItem(attr1.LocalName, attr1.NamespaceURI);
                else
                    attr2 = (XmlAttribute)element2.Attributes.GetNamedItem(attr1.Name);

                if (attr2 == null || !StringEquals(attr1.InnerText, attr2.InnerText))
                    //if (attr2 == null || !equals(attr1.getTextContent(), attr2.getTextContent()))
                    return false;
            }

            return true;
        }

        protected static void CompactChildNodesR(XmlElement parentElement, String childTagName)
        {
            XmlNodeList childNodes = parentElement.ChildNodes;
            for (int i = 0; i < childNodes.Count - 1; i++)
            {
                XmlNode child1 = childNodes[i];
                XmlNode child2 = childNodes[i + 1];
                if (!CanBeMerged(child1, child2, childTagName))
                    continue;

                // merge
                while (child2.ChildNodes.Count > 0)
                    child1.AppendChild(child2.FirstChild);
                child2.ParentNode.RemoveChild(child2);
                i--;
            }

            childNodes = parentElement.ChildNodes;
            for (int i = 0; i < childNodes.Count - 1; i++)
            {
                XmlNode child = childNodes[i];
                if (child is XmlElement)
                {
                    CompactChildNodesR((XmlElement)child, childTagName);
                }
            }
        }
        static bool StringEquals(String str1, String str2)
        {
            return str1 == null ? str2 == null : str1.Equals(str2);
        }
        public static String GetBorderType(BorderCode borderCode)
        {
            if (borderCode == null)
                throw new ArgumentNullException("borderCode is null");

            switch (borderCode.BorderType)
            {
                case 1:
                case 2:
                    return "solid";
                case 3:
                    return "double";
                case 5:
                    return "solid";
                case 6:
                    return "dotted";
                case 7:
                case 8:
                    return "dashed";
                case 9:
                    return "dotted";
                case 10:
                case 11:
                case 12:
                case 13:
                case 14:
                case 15:
                case 16:
                case 17:
                case 18:
                case 19:
                    return "double";
                case 20:
                    return "solid";
                case 21:
                    return "double";
                case 22:
                    return "dashed";
                case 23:
                    return "dashed";
                case 24:
                    return "ridge";
                case 25:
                    return "grooved";
                default:
                    return "solid";
            }
        }

        public static String GetBorderWidth(BorderCode borderCode)
        {
            int lineWidth = borderCode.LineWidth;
            int pt = lineWidth / 8;
            int pte = lineWidth - pt * 8;

            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(pt);
            stringBuilder.Append(".");
            stringBuilder.Append(1000 / 8 * pte);
            stringBuilder.Append("pt");
            return stringBuilder.ToString();
        }

        public static String GetBulletText(ListTables listTables,
            Paragraph paragraph, int listId)
        {
            ListLevel listLevel = listTables.GetLevel(listId,
                    paragraph.GetIlvl());

            if (listLevel.GetNumberText() == null)
                return string.Empty;

            StringBuilder bulletBuffer = new StringBuilder();
            char[] xst = listLevel.GetNumberText().ToCharArray();
            foreach (char element in xst)
            {
                if (element < 9)//todo:review_antony
                {
                    ListLevel numLevel = listTables.GetLevel(listId, element);

                    int num = numLevel.GetStartAt();
                    bulletBuffer.Append(NumberFormatter.GetNumber(num,
                            listLevel.GetNumberFormat()));

                    if (numLevel == listLevel)
                    {
                        numLevel.SetStartAt(numLevel.GetStartAt() + 1);
                    }

                }
                else
                {
                    bulletBuffer.Append(element);
                }
            }

            byte follow = listLevel.GetTypeOfCharFollowingTheNumber();
            switch (follow)
            {
                case 0:
                    bulletBuffer.Append("\t");
                    break;
                case 1:
                    bulletBuffer.Append(" ");
                    break;
                default:
                    break;
            }

            return bulletBuffer.ToString();
        }

        public static String GetColor(int ico)
        {
            switch (ico)
            {
                case 1:
                    return "black";
                case 2:
                    return "blue";
                case 3:
                    return "cyan";
                case 4:
                    return "green";
                case 5:
                    return "magenta";
                case 6:
                    return "red";
                case 7:
                    return "yellow";
                case 8:
                    return "white";
                case 9:
                    return "darkblue";
                case 10:
                    return "darkcyan";
                case 11:
                    return "darkgreen";
                case 12:
                    return "darkmagenta";
                case 13:
                    return "darkred";
                case 14:
                    return "darkyellow";
                case 15:
                    return "darkgray";
                case 16:
                    return "lightgray";
                default:
                    return "black";
            }
        }

        public static String GetOpacity(int argbValue)
        {
            int opacity = (int)((argbValue & 0xFF000000L) >> 24);// todo: review code, java use operater >>>
            if (opacity == 0 || opacity == 0xFF)
                return ".0";

            return "" + (opacity / (float)0xFF);
        }

        public static String GetColor24(int argbValue)
        {
            if (argbValue == -1)
                throw new ArgumentException("This colorref is empty");

            int bgrValue = argbValue & 0x00FFFFFF;
            int rgbValue = (bgrValue & 0x0000FF) << 16 | (bgrValue & 0x00FF00)
                    | (bgrValue & 0xFF0000) >> 16;

            // http://www.w3.org/TR/REC-html40/types.html#h-6.5
            switch (rgbValue)
            {
                case 0xFFFFFF:
                    return "white";
                case 0xC0C0C0:
                    return "silver";
                case 0x808080:
                    return "gray";
                case 0x000000:
                    return "black";
                case 0xFF0000:
                    return "red";
                case 0x800000:
                    return "maroon";
                case 0xFFFF00:
                    return "yellow";
                case 0x808000:
                    return "olive";
                case 0x00FF00:
                    return "lime";
                case 0x008000:
                    return "green";
                case 0x00FFFF:
                    return "aqua";
                case 0x008080:
                    return "teal";
                case 0x0000FF:
                    return "blue";
                case 0x000080:
                    return "navy";
                case 0xFF00FF:
                    return "fuchsia";
                case 0x800080:
                    return "purple";
            }

            StringBuilder result = new StringBuilder("#");
            String hex = rgbValue.ToString("x");
            for (int i = hex.Length; i < 6; i++)
            {
                result.Append('0');
            }
            result.Append(hex);
            return result.ToString();
        }

        public static String GetJustification(int js)
        {
            switch (js)
            {
                case 0:
                    return "start";
                case 1:
                    return "center";
                case 2:
                    return "end";
                case 3:
                case 4:
                    return "justify";
                case 5:
                    return "center";
                case 6:
                    return "left";
                case 7:
                    return "start";
                case 8:
                    return "end";
                case 9:
                    return "justify";
            }
            return "";
        }

        public static String GetLanguage(int languageCode)
        {
            switch (languageCode)
            {
                case 1024:
                    return string.Empty;
                case 1033:
                    return "en-us";
                case 1049:
                    return "ru-ru";
                case 2057:
                    return "en-uk";
                default:
                    logger.Log(POILogger.WARN, "Uknown or unmapped language code: ", languageCode);
                    return string.Empty;
            }
        }

        public static String GetListItemNumberLabel(int number, int format)
        {

            if (format != 0)
                System.Console.WriteLine("NYI: toListItemNumberLabel(): " + format);

            return number.ToString();
        }

        public static HWPFDocumentCore LoadDoc(DirectoryNode root)
        {
            try
            {
                return new HWPFDocument(root);
            }
            catch (OldWordFileFormatException exc)
            {
                return new HWPFOldDocument(root);
            }
        }
        public static HWPFDocumentCore LoadDoc(POIFSFileSystem poifsFileSystem)
        {
            return LoadDoc( poifsFileSystem.Root );
        }
        public static HWPFDocumentCore LoadDoc(Stream inputStream)
        {
            return LoadDoc(HWPFDocumentCore.VerifyAndBuildPOIFS(inputStream));
        }
        public static HWPFDocumentCore LoadDoc(string docFile)
        {
            FileStream istream = new FileStream(docFile, FileMode.Open);
            try
            {
                return LoadDoc(istream);
            }
            finally
            {
                if (istream != null)
                    istream.Close();
                istream = null;
            }
        }

        public static String SubstringBeforeLast(String str, String separator)
        {
            if (string.IsNullOrEmpty(str) || string.IsNullOrEmpty(separator))
            {
                return str;
            }
            int pos = str.LastIndexOf(separator);
            if (pos == -1)
            {
                return str;
            }
            return str.Substring(0, pos);
        }

    }
}
