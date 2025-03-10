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

namespace NPOI.XSSF.UserModel
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula.UDF;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.Streaming;

    /**
     * Internal POI use only - parent of XSSF and SXSSF Evaluation workbooks
     */
    public abstract class BaseXSSFEvaluationWorkbook : IFormulaRenderingWorkbook, IEvaluationWorkbook, IFormulaParsingWorkbook
    {
        protected XSSFWorkbook _uBook;

        public virtual void ClearAllCachedResultValues()
        {
            _tableCache = null;
        }

        protected BaseXSSFEvaluationWorkbook(XSSFWorkbook book)
        {
            _uBook = book;
        }

        private int ConvertFromExternalSheetIndex(int externSheetIndex)
        {
            return externSheetIndex;
        }
        /**
         * XSSF doesn't use external sheet indexes, so when asked treat
         * it just as a local index
         */
        public int ConvertFromExternSheetIndex(int externSheetIndex)
        {
            return externSheetIndex;
        }
        /**
         * @return  the external sheet index of the sheet with the given internal
         * index. Used by some of the more obscure formula and named range things.
         * Fairly easy on XSSF (we think...) since the internal and external
         * indices are the same
         */
        private int ConvertToExternalSheetIndex(int sheetIndex)
        {
            return sheetIndex;
        }

        public int GetExternalSheetIndex(String sheetName)
        {
            int sheetIndex = _uBook.GetSheetIndex(sheetName);
            return ConvertToExternalSheetIndex(sheetIndex);
        }

        private int ResolveBookIndex(String bookName)
        {
            // Strip the [] wrapper, if still present
            if (bookName.StartsWith("[") && bookName.EndsWith("]"))
            {
                bookName = bookName.Substring(1, bookName.Length - 2);
            }

            // Is it already in numeric form?
            try
            {
                return Int32.Parse(bookName);
            }
            catch (FormatException ) { }

            // Look up an External Link Table for this name
            List<ExternalLinksTable> tables = _uBook.ExternalLinksTable;
            int index = FindExternalLinkIndex(bookName, tables);
            if (index != -1) return index;

            // Is it an absolute file reference?
            if (bookName.StartsWith("'file:///") && bookName.EndsWith("'"))
            {
                String relBookName = bookName.Substring(bookName.LastIndexOf('/') + 1);
                relBookName = relBookName.Substring(0, relBookName.Length - 1); // Trailing '

                // Try with this name
                index = FindExternalLinkIndex(relBookName, tables);
                if (index != -1) return index;

                // If we Get here, it's got no associated proper links yet
                // So, add the missing reference and return
                // Note - this is really rather nasty...
                ExternalLinksTable fakeLinkTable = new FakeExternalLinksTable(relBookName);
                tables.Add(fakeLinkTable);
                return tables.Count; // 1 based results, 0 = current workbook
            }

            // Not properly referenced
            throw new Exception("Book not linked for filename " + bookName);
        }
        /* case-sensitive */
        private int FindExternalLinkIndex(String bookName, List<ExternalLinksTable> tables)
        {
            int i = 0;
            foreach (ExternalLinksTable table in tables)
            {
                if (table.LinkedFileName.Equals(bookName))
                {
                    return i + 1; // 1 based results, 0 = current workbook
                }
                i++;
            }
            return -1;
        }
        private sealed class FakeExternalLinksTable : ExternalLinksTable
        {
            private readonly String fileName;
            internal FakeExternalLinksTable(string fileName)
            {
                this.fileName = fileName;
            }
            public override string LinkedFileName
            {
                get
                {
                    return fileName;
                }
            }
        }
        /// <summary>
        /// Return EvaluationName wrapper around the matching XSSFName (named range)
        /// </summary>
        /// <param name="name">case-aware but case-insensitive named range in workbook</param>
        /// <param name="sheetIndex">index of sheet if named range scope is limited to one sheet
        ///   if named range scope is global to the workbook, sheetIndex is -1.</param>
        /// <returns>If name is a named range in the workbook, returns
        /// EvaluationName corresponding to that named range 
        /// Returns null if there is no named range with the same name and scope in the workbook
        /// </returns>
        public IEvaluationName GetName(String name, int sheetIndex)
        {
            for (int i = 0; i < _uBook.NumberOfNames; i++)
            {
                XSSFName nm = _uBook.GetNameAt(i) as XSSFName;
                String nameText = nm.NameName;
                int nameSheetindex = nm.SheetIndex;
                if (name.Equals(nameText, StringComparison.CurrentCultureIgnoreCase) &&
                       (nameSheetindex == -1 || nameSheetindex == sheetIndex))
                {
                    return new Name(nm, i, this);
                }
            }
            return sheetIndex == -1 ? null : GetName(name, -1);
        }

        public String GetSheetName(int sheetIndex)
        {
            return _uBook.GetSheetName(sheetIndex);
        }

        public ExternalName GetExternalName(int externSheetIndex, int externNameIndex)
        {
            throw new InvalidOperationException("HSSF-style external references are not supported for XSSF");
        }

        public ExternalName GetExternalName(String nameName, String sheetName, int externalWorkbookNumber)
        {
            if (externalWorkbookNumber > 0)
            {
                // External reference - reference is 1 based, link table is 0 based
                int linkNumber = externalWorkbookNumber - 1;
                ExternalLinksTable linkTable = _uBook.ExternalLinksTable[(linkNumber)];

                foreach (IName name in linkTable.DefinedNames)
                {
                    if (name.NameName.Equals(nameName))
                    {
                        // HSSF returns one sheet higher than normal, and various bits
                        //  of the code assume that. So, make us match that behaviour!
                        int nameSheetIndex = name.SheetIndex + 1;

                        // TODO Return a more specialised form of this, see bug #56752
                        // Should include the cached values, for in case that book isn't available
                        // Should support XSSF stuff Lookups
                        return new ExternalName(nameName, -1, nameSheetIndex);
                    }
                }
                throw new ArgumentException("Name '" + nameName + "' not found in " +
                                                   "reference to " + linkTable.LinkedFileName);
            }
            else
            {
                // Internal reference
                int nameIdx = _uBook.GetNameIndex(nameName);
                return new ExternalName(nameName, nameIdx, 0);  // TODO Is this right?
            }

        }

        /// <summary>
        /// Return an external name (named range, function, user-defined function) Pxg
        /// </summary>
        /// <param name="name"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public Ptg GetNameXPtg(String name, SheetIdentifier sheet)
        {
            // First, try to find it as a User Defined Function
            IndexedUDFFinder udfFinder = (IndexedUDFFinder)GetUDFFinder();
            FreeRefFunction func = udfFinder.FindFunction(name);
            if (func != null)
            {
                return new NameXPxg(null, name);
            }

            // Otherwise, try it as a named range
            if (sheet == null)
            {
                if (_uBook.GetNames(name).Count > 0)
                {
                    return new NameXPxg(null, name);
                }
                return null;
            }
            if (sheet._sheetIdentifier == null)
            {
                // Workbook + Named Range only
                int bookIndex = ResolveBookIndex(sheet._bookName);
                return new NameXPxg(bookIndex, null, name);
            }

            // Use the sheetname and process
            String sheetName = sheet._sheetIdentifier.Name;

            if (sheet._bookName != null)
            {
                int bookIndex = ResolveBookIndex(sheet._bookName);
                return new NameXPxg(bookIndex, sheetName, name);
            }
            else
            {
                return new NameXPxg(sheetName, name);
            }
        }
        public Ptg Get3DReferencePtg(CellReference cell, SheetIdentifier sheet)
        {
            if (sheet._bookName != null)
            {
                int bookIndex = ResolveBookIndex(sheet._bookName);
                return new Ref3DPxg(bookIndex, sheet, cell);
            }
            else
            {
                return new Ref3DPxg(sheet, cell);
            }
        }
        public Ptg Get3DReferencePtg(AreaReference area, SheetIdentifier sheet)
        {
            if (sheet._bookName != null)
            {
                int bookIndex = ResolveBookIndex(sheet._bookName);
                return new Area3DPxg(bookIndex, sheet, area);
            }
            else
            {
                return new Area3DPxg(sheet, area);
            }
        }

        public String ResolveNameXText(NameXPtg n)
        {
            int idx = n.NameIndex;
            String name = null;

            // First, try to find it as a User Defined Function
            IndexedUDFFinder udfFinder = (IndexedUDFFinder)GetUDFFinder();
            name = udfFinder.GetFunctionName(idx);
            if (name != null) return name;

            // Otherwise, try it as a named range
            XSSFName xname = _uBook.GetNameAt(idx) as XSSFName;
            if (xname != null)
            {
                name = xname.NameName;
            }

            return name;
        }

        public ExternalSheet GetExternalSheet(int externSheetIndex)
        {
            throw new InvalidOperationException("HSSF-style external references are not supported for XSSF");
        }
        public ExternalSheet GetExternalSheet(String firstSheetName, String lastSheetName, int externalWorkbookNumber)
        {
            String workbookName;
            if (externalWorkbookNumber > 0)
            {
                // External reference - reference is 1 based, link table is 0 based
                int linkNumber = externalWorkbookNumber - 1;
                ExternalLinksTable linkTable = _uBook.ExternalLinksTable[(linkNumber)];
                workbookName = linkTable.LinkedFileName;
            }
            else
            {
                // Internal reference
                workbookName = null;
            }

            if (lastSheetName == null || firstSheetName.Equals(lastSheetName))
            {
                return new ExternalSheet(workbookName, firstSheetName);
            }
            else
            {
                return new ExternalSheetRange(workbookName, firstSheetName, lastSheetName);
            }
        }

        public int GetExternalSheetIndex(String workbookName, String sheetName)
        {
            throw new Exception("not implemented yet");
        }
        public int GetSheetIndex(String sheetName)
        {
            return _uBook.GetSheetIndex(sheetName);
        }

        public String GetSheetFirstNameByExternSheet(int externSheetIndex)
        {
            int sheetIndex = ConvertFromExternalSheetIndex(externSheetIndex);
            return _uBook.GetSheetName(sheetIndex);
        }
        public String GetSheetLastNameByExternSheet(int externSheetIndex)
        {
            // XSSF does multi-sheet references differently, so this is the same as the first
            return GetSheetFirstNameByExternSheet(externSheetIndex);
        }

        public String GetNameText(NamePtg namePtg)
        {
            return _uBook.GetNameAt(namePtg.Index).NameName;
        }
        public IEvaluationName GetName(NamePtg namePtg)
        {
            int ix = namePtg.Index;
            return new Name(_uBook.GetNameAt(ix) as XSSFName, ix, this);
        }
        public IName CreateName()
        {
            return _uBook.CreateName();
        }


        /*
         * TODO: data tables are stored at the workbook level in XSSF, but are bound to a single sheet.
         *       The current code structure has them hanging off XSSFSheet, but formulas reference them
         *       only by name (names are global, and case insensitive).
         *       This map stores names as lower case for case-insensitive lookups.
         *
         * FIXME: Caching tables by name here for fast formula lookup means the map is out of date if
         *       a table is renamed or added/removed to a sheet after the map is created.
         *
         *       Perhaps tables can be managed similar to PivotTable references above?
         */
        private Dictionary<String, XSSFTable> _tableCache = null;
        private Dictionary<String, XSSFTable> GetTableCache()
        {
            if (_tableCache != null)
            {
                return _tableCache;
            }
            // FIXME: use org.apache.commons.collections.map.CaseInsensitiveMap
            _tableCache = new Dictionary<String, XSSFTable>();

            foreach (ISheet sheet in _uBook)
            {
                foreach (XSSFTable tbl in ((XSSFSheet)sheet).GetTables())
                {
                    String lname = tbl.Name.ToLower(CultureInfo.CurrentCulture);
                    _tableCache.Add(lname, tbl);
                }
            }
            return _tableCache;
        }

        /**
         * Returns the data table with the given name (case insensitive).
         * Tables are cached for performance (formula evaluation looks them up by name repeatedly).
         * After the first table lookup, adding or removing a table from the document structure will cause trouble.
         * This is meant to be used on documents whose structure is essentially static at the point formulas are evaluated.
         * 
         * @param name the data table name (case-insensitive)
         * @return The Data table in the workbook named <tt>name</tt>, or <tt>null</tt> if no table is named <tt>name</tt>.
         * @since 3.15 beta 2
         */
        public ITable GetTable(String name)
        {
            if (name == null) return null;
            String lname = name.ToLower(CultureInfo.CurrentCulture);
            return GetTableCache()[lname];
        }

        public UDFFinder GetUDFFinder()
        {
            return _uBook.GetUDFFinder();
        }

        private sealed class Name : IEvaluationName
        {

            private readonly XSSFName _nameRecord;
            private readonly int _index;
            private readonly IFormulaParsingWorkbook _fpBook;

            public Name(XSSFName name, int index, IFormulaParsingWorkbook fpBook)
            {
                _nameRecord = name;
                _index = index;
                _fpBook = fpBook;
            }

            public Ptg[] NameDefinition
            {
                get
                {
                    return FormulaParser.Parse(_nameRecord.RefersToFormula, _fpBook, FormulaType.NamedRange, _nameRecord.SheetIndex);
                }
                
            }

            public String NameText
            {
                get
                {
                    return _nameRecord.NameName;
                }
                
            }

            public bool HasFormula
            {
                get
                {
                    // TODO - no idea if this is right
                    CT_DefinedName ctn = _nameRecord.GetCTName();
                    String strVal = ctn.Value;
                    return !ctn.function && strVal != null && strVal.Length > 0;
                }
                
            }

            public bool IsFunctionName
            {
                get
                {
                    return _nameRecord.IsFunctionName;
                }
            }

            public bool IsRange
            {
                get
                {
                    return HasFormula; // TODO - is this right?
                }
            }
            public NamePtg CreatePtg()
            {
                return new NamePtg(_index);
            }
        }

        public SpreadsheetVersion GetSpreadsheetVersion()
        {
            return SpreadsheetVersion.EXCEL2007;
        }
        
        public abstract int GetSheetIndex(IEvaluationSheet sheet);

        public abstract IEvaluationSheet GetSheet(int sheetIndex);

        public abstract Ptg[] GetFormulaTokens(IEvaluationCell cell);
    }

}