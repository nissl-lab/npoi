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
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula.Udf;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.Util;
using NPOI.SS.Util;
using System.Collections.Generic;

namespace NPOI.XSSF.UserModel
{
    /**
     * Internal POI use only
     *
     * @author Josh Micich
     */
    public class XSSFEvaluationWorkbook : IFormulaRenderingWorkbook, 
        IEvaluationWorkbook, IFormulaParsingWorkbook
    {

        private XSSFWorkbook _uBook;

        public static XSSFEvaluationWorkbook Create(IWorkbook book)
        {
            if (book == null)
            {
                return null;
            }
            return new XSSFEvaluationWorkbook(book);
        }

        protected XSSFEvaluationWorkbook(IWorkbook book)
        {
            _uBook = (XSSFWorkbook)book;
        }

        private int ConvertFromExternalSheetIndex(int externSheetIndex)
        {
            return externSheetIndex;
        }
        /**
         * @return the sheet index of the sheet with the given external index.
         */
        public int ConvertFromExternSheetIndex(int externSheetIndex)
        {
            return externSheetIndex;
        }
        /**
         * @return  the external sheet index of the sheet with the given internal
         * index. Used by some of the more obscure formula and named range things.
         * Fairly easy on XSSF (we think...) since the internal and external
         * indicies are the same
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
            catch (FormatException e) { }

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
        private int FindExternalLinkIndex(String bookName, List<ExternalLinksTable> tables)
        {
            for (int i = 0; i < tables.Count; i++)
            {
                if (tables[(i)].LinkedFileName.Equals(bookName))
                {
                    return i + 1; // 1 based results, 0 = current workbook
                }
            }
            return -1;
        }
        private class FakeExternalLinksTable : ExternalLinksTable
        {
            private String fileName;
            public FakeExternalLinksTable(String fileName)
            {
                this.fileName = fileName;
            }
            public override String LinkedFileName
            {
                get
                {
                    return fileName;
                }
                set
                {
                    this.fileName = value;
                }
            }
        }


        public IEvaluationName GetName(String name, int sheetIndex)
        {
            for (int i = 0; i < _uBook.NumberOfNames; i++)
            {
                IName nm = _uBook.GetNameAt(i);
                String nameText = nm.NameName;
                int nameSheetindex = nm.SheetIndex;
                if (name.Equals(nameText, StringComparison.CurrentCultureIgnoreCase) &&
                       (nameSheetindex == -1 || nameSheetindex == sheetIndex))
                {
                    return new Name(_uBook.GetNameAt(i), i, this);
                }
            }
            return sheetIndex == -1 ? null : GetName(name, -1);
        }

        public int GetSheetIndex(IEvaluationSheet EvalSheet)
        {
            XSSFSheet sheet = ((XSSFEvaluationSheet)EvalSheet).GetXSSFSheet();
            return _uBook.GetSheetIndex(sheet);
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
                ExternalLinksTable linkTable = _uBook.ExternalLinksTable[linkNumber];
                foreach (IName name in linkTable.DefinedNames)
                {
                    if (name.NameName.Equals(nameName))
                    {
                        // HSSF returns one sheet higher than normal, and various bits
                        //  of the code assume that. So, make us match that behaviour!
                        int nameSheetIndex = name.SheetIndex + 1;
                        // TODO Return a more specialised form of this, see bug #56752
                        // Should include the cached values, for in case that book isn't available
                        // Should support XSSF stuff lookups
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
        public Ptg GetNameXPtg(String name, SheetIdentifier sheet)
        {
            IndexedUDFFinder udfFinder = (IndexedUDFFinder)GetUDFFinder();
            FreeRefFunction func = udfFinder.FindFunction(name);
            if (func != null)
            {
                return new NameXPxg(null, name);
            }

            // Otherwise, try it as a named range
            if (sheet == null)
            {
                if (_uBook.GetNameIndex(name) > -1)
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
            IName xname = _uBook.GetNameAt(idx);
            if (xname != null)
            {
                name = xname.NameName;
            }

            return name;
        }

        public IEvaluationSheet GetSheet(int sheetIndex)
        {
            return new XSSFEvaluationSheet(_uBook.GetSheetAt(sheetIndex));
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
            throw new RuntimeException("not implemented yet");
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
            return new Name(_uBook.GetNameAt(ix), ix, this);
        }
        public Ptg[] GetFormulaTokens(IEvaluationCell EvalCell)
        {
            XSSFCell cell = ((XSSFEvaluationCell)EvalCell).GetXSSFCell();
            XSSFEvaluationWorkbook frBook = XSSFEvaluationWorkbook.Create(_uBook);
            
            return FormulaParser.Parse(cell.CellFormula, frBook, FormulaType.Cell, _uBook.GetSheetIndex(cell.Sheet));
        }

        public UDFFinder GetUDFFinder()
        {
            return _uBook.GetUDFFinder();
        }

        /**
         * XSSF allows certain extra textual characters in the formula that
         *  HSSF does not. As these can't be composed down to HSSF-compatible
         *  Ptgs, this method strips them out for us.
         */
        private String CleanXSSFFormulaText(String text)
        {
            // Newlines are allowed in XSSF
            text = text.Replace("\\n", "").Replace("\\r", "");

            // All done with cleaning
            return text;
        }

        private class Name : IEvaluationName
        {

            private XSSFName _nameRecord;
            private int _index;
            private IFormulaParsingWorkbook _fpBook;

            public Name(IName name, int index, IFormulaParsingWorkbook fpBook)
            {
                _nameRecord = (XSSFName)name;
                _index = index;
                _fpBook = fpBook;
            }

            public Ptg[] NameDefinition
            {
                get
                {

                    return FormulaParser.Parse(_nameRecord.RefersToFormula, _fpBook, 
                        FormulaType.NamedRange, _nameRecord.SheetIndex);
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
    }
}

