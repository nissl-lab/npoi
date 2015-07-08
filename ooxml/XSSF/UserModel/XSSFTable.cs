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

using System.Xml;
using System.Collections.Generic;
using NPOI.SS.Util;
using System;
using NPOI.OpenXml4Net.OPC;
using System.IO;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.Util;
using System.Collections;
using NPOI.XSSF.UserModel.Helpers;
namespace NPOI.XSSF.UserModel
{

    /**
     * 
     * This class : the Table Part (Open Office XML Part 4:
     * chapter 3.5.1)
     * 
     * This implementation works under the assumption that a table Contains mappings to a subtree of an XML.
     * The root element of this subtree an occur multiple times (one for each row of the table). The child nodes
     * of the root element can be only attributes or element with maxOccurs=1 property set
     * 
     *
     * @author Roberto Manicardi
     */
    public class XSSFTable : POIXMLDocumentPart
    {

        private CT_Table ctTable;
        private List<XSSFXmlColumnPr> xmlColumnPr;
        private CellReference startCellReference;
        private CellReference endCellReference;
        private String commonXPath;


        public XSSFTable()
            : base()
        {

            ctTable = new CT_Table();

        }

        internal XSSFTable(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
            XmlDocument xml = ConvertStreamToXml(part.GetInputStream());
            ReadFrom(xml);
        }

        public void ReadFrom(XmlDocument xmlDoc)
        {
            try
            {
                TableDocument doc = TableDocument.Parse(xmlDoc, NamespaceManager);
                ctTable = doc.GetTable();
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }

        public XSSFSheet GetXSSFSheet()
        {
            return (XSSFSheet)GetParent();
        }

        public void WriteTo(Stream out1)
        {
            UpdateHeaders();
            TableDocument doc = new TableDocument();
            doc.SetTable(ctTable);
            doc.Save(out1);
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            WriteTo(out1);
            out1.Close();
        }

        public CT_Table GetCTTable()
        {
            return ctTable;
        }

        /**
         * Checks if this Table element Contains even a single mapping to the map identified by id
         * @param id the XSSFMap ID
         * @return true if the Table element contain mappings
         */
        public bool MapsTo(long id)
        {
            bool maps = false;

            List<XSSFXmlColumnPr> pointers = GetXmlColumnPrs();

            foreach (XSSFXmlColumnPr pointer in pointers)
            {
                if (pointer.GetMapId() == id)
                {
                    maps = true;
                    break;
                }
            }

            return maps;
        }


        /**
         * 
         * Calculates the xpath of the root element for the table. This will be the common part
         * of all the mapping's xpaths
         * 
         * @return the xpath of the table's root element
         */
        public String GetCommonXpath()
        {

            if (commonXPath == null)
            {

                Array commonTokens = null;

                foreach (CT_TableColumn column in ctTable.tableColumns.tableColumn)
                {
                    if (column.xmlColumnPr != null)
                    {
                        String xpath = column.xmlColumnPr.xpath;
                        String[] tokens = xpath.Split(new char[] { '/' });
                        if (commonTokens==null)
                        {
                            commonTokens = tokens;
                        }
                        else
                        {
                            int maxLenght = commonTokens.Length > tokens.Length ? tokens.Length : commonTokens.Length;
                            for (int i = 0; i < maxLenght; i++)
                            {
                                if (!commonTokens.GetValue(i).Equals(tokens[i]))
                                {
                                    ArrayList subCommonTokens = Arrays.AsList(commonTokens).GetRange(0, i);
                                    commonTokens = subCommonTokens.ToArray(typeof(string));
                                    break;


                                }
                            }
                        }

                    }
                }


                commonXPath = "";

                for (int i = 1; i < commonTokens.Length; i++)
                {
                    commonXPath += "/" + commonTokens.GetValue(i);

                }
            }

            return commonXPath;
        }


        public List<XSSFXmlColumnPr> GetXmlColumnPrs()
        {

            if (xmlColumnPr == null)
            {
                xmlColumnPr = new List<XSSFXmlColumnPr>();
                foreach (CT_TableColumn column in ctTable.tableColumns.tableColumn)
                {
                    if (column.xmlColumnPr != null)
                    {
                        XSSFXmlColumnPr columnPr = new XSSFXmlColumnPr(this, column, column.xmlColumnPr);
                        xmlColumnPr.Add(columnPr);
                    }
                }
            }
            return xmlColumnPr;
        }

        /**
         * @return the name of the Table, if set
         */
        public String Name
        {
            get
            {
                return ctTable.name;
            }
            set 
            {
                ctTable.name = value;
            }
        }

        /**
         * @return the display name of the Table, if set
         */
        public String DisplayName
        {
            get
            {
                return ctTable.displayName;
            }
            set
            {
                ctTable.displayName = value;
            }

        }

        /**
         * @return  the number of mapped table columns (see Open Office XML Part 4: chapter 3.5.1.4)
         */
        public long NumberOfMappedColumns
        {
            get
            {
                return ctTable.tableColumns.count;
            }
        }


        /**
         * @return The reference for the cell in the top-left part of the table
         * (see Open Office XML Part 4: chapter 3.5.1.2, attribute ref) 
         *
         */
        public CellReference GetStartCellReference()
        {

            if (startCellReference == null)
            {
                String ref1 = ctTable.@ref;
                if(ref1 != null) {
                    String[] boundaries = ref1.Split(":".ToCharArray());
                    String from = boundaries[0];
                    startCellReference = new CellReference(from);
                }
            }
            return startCellReference;
        }

        /**
         * @return The reference for the cell in the bottom-right part of the table
         * (see Open Office XML Part 4: chapter 3.5.1.2, attribute ref)
         *
         */
        public CellReference GetEndCellReference()
        {

            if (endCellReference == null)
            {

                String ref1 = ctTable.@ref;
                String[] boundaries = ref1.Split(new char[] { ':' });
                String from = boundaries[1];
                endCellReference = new CellReference(from);
            }
            return endCellReference;
        }


        /**
         *  @return the total number of rows in the selection. (Note: in this version autofiltering is ignored)
         *
         */
        public int RowCount
        {
            get
            {
                CellReference from = GetStartCellReference();
                CellReference to = GetEndCellReference();

                int rowCount = -1;
                if (from != null && to != null)
                {
                    rowCount = to.Row - from.Row;
                }
                return rowCount;
            }
        }


        /**
     * Synchronize table headers with cell values in the parent sheet.
     * Headers <em>must</em> be in sync, otherwise Excel will display a
     * "Found unreadable content" message on startup.
     */
        public void UpdateHeaders()
        {
            XSSFSheet sheet = (XSSFSheet)GetParent();
            CellReference ref1 = GetStartCellReference() as CellReference;
            if (ref1 == null) return;

            int headerRow = ref1.Row;
            int firstHeaderColumn = ref1.Col;
            XSSFRow row = sheet.GetRow(headerRow) as XSSFRow;

            if (row != null && row.GetCTRow() != null)
            {
                int cellnum = firstHeaderColumn;
                foreach (CT_TableColumn col in GetCTTable().tableColumns.tableColumn)
                {
                    XSSFCell cell = row.GetCell(cellnum) as XSSFCell;
                    if (cell != null)
                    {
                        col.name = (cell.StringCellValue);
                    }
                    cellnum++;
                }
            }
        }
    }
}