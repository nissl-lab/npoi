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
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
using System.Globalization;
using NPOI.SS;
using System.Linq;
using NPOI.OOXML.XSSF.UserModel;

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
    public class XSSFTable : POIXMLDocumentPart, ITable
    {
        private CT_Table ctTable;
        private List<XSSFXmlColumnPr> xmlColumnPrs;
        private List<XSSFTableColumn> tableColumns;
        private Dictionary<String, int> columnMap;
        private CellReference startCellReference;
        private CellReference endCellReference;
        private String commonXPath;


        public XSSFTable()
            : base()
        {

            ctTable = new CT_Table();

        }

        internal XSSFTable(PackagePart part)
            : base(part)
        {
            XmlDocument xml = ConvertStreamToXml(part.GetInputStream());
            ReadFrom(xml);
        }

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        protected XSSFTable(PackagePart part, PackageRelationship rel)
            : this(part)
        {

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
                if (pointer.MapId == id)
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

                foreach (XSSFTableColumn column in GetColumns())
                {
                    if (column.GetXmlColumnPr() != null)
                    {
                        String xpath = column.GetXmlColumnPr().XPath;
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

        /**
         * Note this list is static - once read, it does not notice later changes to the underlying column structures
         * @return List of XSSFXmlColumnPr
         */
        [Obsolete]
        public List<XSSFXmlColumnPr> GetXmlColumnPrs()
        {

            if (xmlColumnPrs == null)
            {
                xmlColumnPrs = new List<XSSFXmlColumnPr>();
                foreach (CT_TableColumn column in ctTable.tableColumns.tableColumn)
                {
                    if (column.xmlColumnPr != null)
                    {
                        XSSFXmlColumnPr columnPr = new XSSFXmlColumnPr(this, column, column.xmlColumnPr);
                        xmlColumnPrs.Add(columnPr);
                    }
                }
            }
            return xmlColumnPrs;
        }
        private string name;
        /**
         * @return the name of the Table, if set
         */
        public String Name
        {
            get
            {
                if (name == null)
                {
                    Name = ctTable.name;
                }
                return name;
            }
            set 
            {
                if (value == null)
                {
                    ctTable.name=null;
                    name = null;
                    return;
                }
                ctTable.name = value;
                name = value;
            }
        }
        public XSSFTableColumn CreateColumn(String columnName)
        {
            return CreateColumn(columnName, this.ColumnCount);
        }
        public XSSFTableColumn CreateColumn(String columnName, int columnIndex)
        {

            int columnCount = ColumnCount;
            if (columnIndex < 0 || columnIndex > columnCount)
            {
                throw new ArgumentException("Column index out of bounds");
            }

            // Ensure we have Table Columns
            CT_TableColumns columns = ctTable.tableColumns;
            if (columns == null)
            {
                columns = ctTable.AddNewTableColumns();
            }

            // check if name is unique and calculate unique column id 
            long nextColumnId = 0;
            foreach (XSSFTableColumn tableColumn in this.GetColumns())
            {
                if (columnName != null && columnName.Equals(tableColumn.Name,StringComparison.InvariantCultureIgnoreCase))
                {
                    throw new ArgumentException("Column '" + columnName
                            + "' already exists. Column names must be unique per table.");
                }
                nextColumnId = Math.Max(nextColumnId, tableColumn.Id);
            }
            // Bug #62740, the logic was just re-using the existing max ID, not incrementing beyond it.
            nextColumnId++;

            // Add the new Column
            CT_TableColumn column = columns.InsertNewTableColumn(columnIndex);
            columns.count = columns.count;

            column.id = (uint)nextColumnId;
            if (columnName != null)
            {
                column.name = columnName;
            }
            else
            {
                column.name =  "Column " + nextColumnId;
            }

            /*if (ctTable.@ref != null)
            {
                // calculate new area
                int newColumnCount = columnCount + 1;
                CellReference tableStart = StartCellReference;
                CellReference tableEnd = EndCellReference;
                SpreadsheetVersion version = GetXSSFSheet().GetWorkbook().SpreadsheetVersion;
                CellReference newTableEnd = new CellReference(tableEnd.Row,
                        tableStart.Col + newColumnCount - 1);
                AreaReference newTableArea = new AreaReference(tableStart, newTableEnd, version);

                SetCellRef(newTableArea);
            }*/

            UpdateHeaders();

            return GetColumns()[columnIndex];
        }
        /**
         * Get the area reference for the cells which this table covers. The area
         * includes header rows and totals rows.
         *
         * Does not track updates to underlying changes to CTTable To synchronize
         * with changes to the underlying CTTable, call {@link #updateReferences()}.
         * 
         * @return the area of the table
         * @see "Open Office XML Part 4: chapter 3.5.1.2, attribute ref"
         */
        public AreaReference GetCellReferences()
        {
            return new AreaReference(
                    StartCellReference,
                    EndCellReference,
                    SpreadsheetVersion.EXCEL2007
            );
        }
        /**
         * Set the area reference for the cells which this table covers. The area
         * includes includes header rows and totals rows. Automatically synchronizes
         * any changes by calling {@link #updateHeaders()}.
         * 
         * Note: The area's width should be identical to the amount of columns in
         * the table or the table may be invalid. All header rows, totals rows and
         * at least one data row must fit inside the area. Updating the area with
         * this method does not create or remove any columns and does not change any
         * cell values.
         * 
         * @see "Open Office XML Part 4: chapter 3.5.1.2, attribute ref"
         */
        public void SetCellReferences(AreaReference refs)
        {
            SetCellRef(refs);
        }
        protected void SetCellRef(AreaReference refs)
        {

            // Strip the sheet name,
            // CTWorksheet.getTableParts defines in which sheet the table is
            String reference = refs.FormatAsString();
            if (reference.IndexOf('!') != -1)
            {
                reference = reference.Substring(reference.IndexOf('!') + 1);
            }

            // Update
            ctTable.@ref = reference;
            if (ctTable.IsSetAutoFilter)
            {
                String filterRef;
                int totalsRowCount = TotalsRowCount;
                if (totalsRowCount == 0)
                {
                    filterRef = reference;
                }
                else
                {
                    CellReference start = new CellReference(refs.FirstCell.Row, refs.FirstCell.Col);
                    // account for footer row(s) in auto-filter range, which doesn't include footers
                    CellReference end = new CellReference(refs.LastCell.Row - totalsRowCount, refs.LastCell.Col);
                    // this won't have sheet references because we built the cell references without them
                    filterRef = new AreaReference(start, end, SpreadsheetVersion.EXCEL2007).FormatAsString();
                }
                ctTable.autoFilter.@ref =filterRef;
            }

            // Have everything recomputed
            UpdateReferences();
            UpdateHeaders();
        }
        private String styleName;
        public string StyleName
        {
            get {
                if (styleName == null && ctTable.IsSetTableStyleInfo())
                {
                    StyleName = ctTable.tableStyleInfo.name;
                }
                return styleName;
            }
            set
            {
                if (value == null)
                {
                    if (ctTable.IsSetTableStyleInfo())
                    {
                        ctTable.tableStyleInfo.name =null;
                    }
                    styleName = null;
                    return;
                }
                if (!ctTable.IsSetTableStyleInfo())
                {
                    ctTable.AddNewTableStyleInfo();
                }
                ctTable.tableStyleInfo.name = value;
                styleName = value;
            }
        }
        public ITableStyleInfo Style
        {
            get
            {
                if (!ctTable.IsSetTableStyleInfo()) return null;
                return new XSSFTableStyleInfo(((XSSFWorkbook)((XSSFSheet)GetParent()).Workbook).GetStylesSource(), ctTable.tableStyleInfo);
            }
        }
        /**
         * @return the display name of the Table, if set
         */
        public string DisplayName
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
        [Obsolete]
        public long NumberOfMappedColumns
        {
            get
            {
                return ctTable.tableColumns.count;
            }
        }
        public int ColumnCount
        {
            get
            {
                CT_TableColumns tableColumns = ctTable.tableColumns;
                if (tableColumns == null)
                {
                    return 0;
                }
                // Casting to int should be safe here - tables larger than the
                // sheet (which holds the actual data of the table) can't exists.
                return (int)tableColumns.tableColumn.Count();
            }
        }
        /// <summary>
        /// 0 for no totals rows, 1 for totals row shown.
        /// Values > 1 are not currently used by Excel up through 2016, and the OOXML spec
        /// doesn't define how they would be implemented.
        /// </summary>
        public int TotalsRowCount
        {
            get { 
                return (int)ctTable.totalsRowCount;
            }
        }
        /// <summary>
        /// 0 for no header rows, 1 for table headers shown.
        /// Values > 1 might be used by Excel for pivot tables?
        /// </summary>
        public int HeaderRowCount
        {
            get 
            {
                return (int)ctTable.headerRowCount;
            }
        }
        /**
         * @return The reference for the cell in the top-left part of the table
         * (see Open Office XML Part 4: chapter 3.5.1.2, attribute ref) 
         *
         * To synchronize with changes to the underlying CTTable,
         * call {@link #updateReferences()}.
         */
        public CellReference StartCellReference
        {
            get
            {
                if (startCellReference == null)
                {
                    SetCellReferences();
                }
                return startCellReference;
            }
            
        }

        /**
         * @return The reference for the cell in the bottom-right part of the table
         * (see Open Office XML Part 4: chapter 3.5.1.2, attribute ref)
         *
         * Does not track updates to underlying changes to CTTable
         * To synchronize with changes to the underlying CTTable,
         * call {@link #updateReferences()}.
         */
        public CellReference EndCellReference
        {
            get
            {
                if (endCellReference == null)
                {
                    SetCellReferences();
                }
                return endCellReference;
            }
            
        }


        /**
      * @since POI 3.15 beta 3
      */
        private void SetCellReferences()
        {
            String ref1 = ctTable.@ref;
            if (ref1 != null) {
                String[] boundaries = ref1.Split(new char[] { ':' }, 2);
                String from = boundaries[0];
                String to = boundaries[1];
                startCellReference = new CellReference(from);
                endCellReference = new CellReference(to);
            }
        }


        /**
         * Clears the cached values set by {@link #getStartCellReference()}
         * and {@link #getEndCellReference()}.
         * The next call to {@link #getStartCellReference()} and
         * {@link #getEndCellReference()} will synchronize the
         * cell references with the underlying <code>CTTable</code>.
         * Thus, {@link #updateReferences()} is inexpensive.
         *
         * @since POI 3.15 beta 3
         */
        public void UpdateReferences()
        {
            startCellReference = null;
            endCellReference = null;
        }

        /**
         * @return the total number of rows in the selection. (Note: in this version autofiltering is ignored)
         * Returns 0 if the start or end cell references are not set.
         *  
         * To synchronize with changes to the underlying CTTable,
         * call {@link #updateReferences()}.
         */
        public int RowCount
        {
            get
            {
                CellReference from = StartCellReference;
                CellReference to = EndCellReference;

                int rowCount = 0;
                if (from != null && to != null)
                {
                    rowCount = to.Row - from.Row + 1;
                }
                return rowCount;
            }
        }


        /**
         * Synchronize table headers with cell values in the parent sheet.
         * Headers <em>must</em> be in sync, otherwise Excel will display a
         * "Found unreadable content" message on startup.
         * 
         * If calling both {@link #updateReferences()} and
         * {@link #updateHeaders()}, {@link #updateReferences()}
         * should be called first.
         */
        public void UpdateHeaders()
        {
            XSSFSheet sheet = (XSSFSheet)GetParent();
            CellReference ref1 = StartCellReference;
            
            if (ref1 == null) return;

            int headerRow = ref1.Row;
            int firstHeaderColumn = ref1.Col;

            XSSFRow row = sheet.GetRow(headerRow) as XSSFRow;

            if (row != null && row.GetCTRow() != null)
            {
                int cellnum = firstHeaderColumn;
                CT_TableColumns tableColumns = GetCTTable().tableColumns;

                if (tableColumns != null)
                {
                    foreach (CT_TableColumn col in tableColumns.tableColumn)
                    {
                        if (row.GetCell(cellnum) is XSSFCell cell)
                        {
                            col.name = cell.StringCellValue;
                        }
                        cellnum++;
                    }
                }
            }
            tableColumns = null;
            columnMap = null;
            xmlColumnPrs = null;
            commonXPath = null;
        }
        /**
         * Gets the relative column index of a column in this table having the header name <code>column</code>.
         * The column index is relative to the left-most column in the table, 0-indexed.
         * Returns <code>-1</code> if <code>column</code> is not a header name in table.
         *
         * Column Header names are case-insensitive
         *
         * Note: this function caches column names for performance. To flush the cache (because columns
         * have been moved or column headers have been changed), {@link #updateHeaders()} must be called.
         *
         * @since 3.15 beta 2
         */
        public int FindColumnIndex(String columnHeader)
        {
            if (columnHeader == null) return -1;
            if (columnMap == null)
            {
                int count = ColumnCount;
                columnMap = new Dictionary<string, int>(count * 3 / 2);

                int i = 0;
                foreach (XSSFTableColumn column in GetColumns())
                {
                    columnMap.Add(column.Name.ToUpper(CultureInfo.CurrentCulture), i);
                    i++;
                }
            }
            // Table column names with special characters need a single quote escape
            // but the escape is not present in the column definition
            int idx = -1;
            string testKey = columnHeader.Replace("'", "").ToUpper(CultureInfo.CurrentCulture);
            if (columnMap.ContainsKey(testKey))
                idx = columnMap[testKey];
            return idx;
        }
        /// <summary>
        /// Note this list is static - once read, it does not notice later changes to the underlying column structures
        /// </summary>
        /// <returns></returns>
        public List<XSSFTableColumn> GetColumns()
        {
            if (tableColumns == null)
            {
                var columns = new List<XSSFTableColumn>();
                CT_TableColumns ctTableColumns = ctTable.tableColumns;
                if (ctTableColumns != null)
                {
                    foreach (CT_TableColumn column in ctTableColumns.GetTableColumnList())
                    {
                        XSSFTableColumn tableColumn = new XSSFTableColumn(this, column);
                        columns.Add(tableColumn);
                    }
                }
                tableColumns = columns;
            }
            return tableColumns;
        }
        public void RemoveColumn(XSSFTableColumn column)
        {
            int columnIndex = GetColumns().IndexOf(column);
            if (columnIndex >= 0)
            {
                ctTable.tableColumns.RemoveTableColumn(columnIndex);
                UpdateReferences();
                UpdateHeaders();
            }
        }
        public String SheetName
        {
            get
            {
                return GetXSSFSheet().SheetName;
            }
        }

        /// <summary>
        /// This is misleading.  The Spec indicates this is true if the totals row
        /// has<b><i>ever</i></b> been shown, not whether or not it is currently displayed.
        /// </summary>
        public bool IsHasTotalsRow
        {
            get
            {
                return ctTable.totalsRowShown;
            }
            set 
            {
                ctTable.totalsRowShown = value;
            }
        }

        public int StartColIndex
        {
            get
            {
                return StartCellReference.Col;
            }
        }

        public int StartRowIndex
        {
            get
            {
                return StartCellReference.Row;
            }
        }

        public int EndColIndex
        {
            get
            {
                return EndCellReference.Col;
            }
        }

        public int EndRowIndex
        {
            get
            {
                return EndCellReference.Row;
            }
        }
        public bool Contains(CellReference cell)
        {
            if (cell == null) return false;
            // check if cell is on the same sheet as the table
            if (! SheetName.Equals(cell.SheetName)) return false;
            // check if the cell is inside the table
            if (cell.Row >= StartRowIndex
                && cell.Row <= EndRowIndex
                && cell.Col >= StartColIndex
                && cell.Col <= EndColIndex)
            {
                return true;
            }
            return false;
        }
    }
}