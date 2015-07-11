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

namespace NPOI.HSSF.UserModel
{
    using System;
    //using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NPOI.SS;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula.Udf;
    using NPOI.SS.UserModel;
using NPOI.Util;
    using NPOI.SS.Util;
   
    /**
     * Internal POI use only
     * 
     * @author Josh Micich
     */
    public class HSSFEvaluationWorkbook : IFormulaRenderingWorkbook, IEvaluationWorkbook, IFormulaParsingWorkbook
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(HSSFEvaluationWorkbook));
        private HSSFWorkbook _uBook;
        private NPOI.HSSF.Model.InternalWorkbook _iBook;

        public static HSSFEvaluationWorkbook Create(NPOI.SS.UserModel.IWorkbook book)
        {
            if (book == null)
            {
                return null;
            }
            return new HSSFEvaluationWorkbook((HSSFWorkbook)book);
        }

        private HSSFEvaluationWorkbook(HSSFWorkbook book)
        {
            _uBook = book;
            _iBook = book.Workbook;
        }

        public int GetExternalSheetIndex(String sheetName)
        {
            int sheetIndex = _uBook.GetSheetIndex(sheetName);
            return _iBook.CheckExternSheet(sheetIndex);
        }
        public int GetExternalSheetIndex(String workbookName, String sheetName)
        {
            return _iBook.GetExternalSheetIndex(workbookName, sheetName);
        }

        public ExternalName GetExternalName(int externSheetIndex, int externNameIndex)
        {
            return _iBook.GetExternalName(externSheetIndex, externNameIndex);
        }

        public ExternalName GetExternalName(String nameName, String sheetName, int externalWorkbookNumber)
        {
            throw new InvalidOperationException("XSSF-style external names are not supported for HSSF");
        }
        public Ptg Get3DReferencePtg(CellReference cr, SheetIdentifier sheet)
        {
            int extIx = GetSheetExtIx(sheet);
            return new Ref3DPtg(cr, extIx);
        }
        public Ptg Get3DReferencePtg(AreaReference areaRef, SheetIdentifier sheet)
        {
            int extIx = GetSheetExtIx(sheet);
            return new Area3DPtg(areaRef, extIx);
        }
        public Ptg GetNameXPtg(String name, SheetIdentifier sheet)
        {
            int sheetRefIndex = GetSheetExtIx(sheet);
            return _iBook.GetNameXPtg(name, sheetRefIndex, _uBook.GetUDFFinder());
        }


        public IEvaluationName GetName(String name,int sheetIndex)
        {
            for (int i = 0; i < _iBook.NumNames; i++)
            {
                NameRecord nr = _iBook.GetNameRecord(i);
                if (nr.SheetNumber == sheetIndex + 1 && name.Equals(nr.NameText, StringComparison.OrdinalIgnoreCase))
                {
                    return new Name(nr, i);
                }
            }
            return sheetIndex == -1 ? null : GetName(name, -1);
        }

        public int GetSheetIndex(IEvaluationSheet evalSheet)
        {
            HSSFSheet sheet = ((HSSFEvaluationSheet)evalSheet).HSSFSheet;
            return _uBook.GetSheetIndex(sheet);
        }
        public int GetSheetIndex(String sheetName)
        {
            return _uBook.GetSheetIndex(sheetName);
        }

        public String GetSheetName(int sheetIndex)
        {
            return _uBook.GetSheetName(sheetIndex);
        }

        public IEvaluationSheet GetSheet(int sheetIndex)
        {
            return new HSSFEvaluationSheet((HSSFSheet)_uBook.GetSheetAt(sheetIndex));
        }
        public int ConvertFromExternSheetIndex(int externSheetIndex)
        {
            // TODO Update this to expose first and last sheet indexes
            return _iBook.GetFirstSheetIndexFromExternSheetIndex(externSheetIndex);
        }

        public ExternalSheet GetExternalSheet(int externSheetIndex)
        {
            ExternalSheet sheet = _iBook.GetExternalSheet(externSheetIndex);
            if (sheet == null)
            {
                // Try to treat it as a local sheet
                int localSheetIndex = ConvertFromExternSheetIndex(externSheetIndex);
                if (localSheetIndex == -1)
                {
                    // The sheet referenced can't be found, sorry
                    return null;
                }
                if (localSheetIndex == -2)
                {
                    // Not actually sheet based at all - is workbook scoped
                    return null;
                }

                // Look up the local sheet
                String sheetName = GetSheetName(localSheetIndex);

                // Is it a single local sheet, or a range?
                int lastLocalSheetIndex = _iBook.GetLastSheetIndexFromExternSheetIndex(externSheetIndex);
                if (lastLocalSheetIndex == localSheetIndex)
                {
                    sheet = new ExternalSheet(null, sheetName);
                }
                else
                {
                    String lastSheetName = GetSheetName(lastLocalSheetIndex);
                    sheet = new ExternalSheetRange(null, sheetName, lastSheetName);
                }
            }
            return sheet;
        }
        public ExternalSheet GetExternalSheet(String firstSheetName, string lastSheetName, int externalWorkbookNumber)
        {
            throw new InvalidOperationException("XSSF-style external references are not supported for HSSF");
        }

        public String ResolveNameXText(NameXPtg n)
        {
            return _iBook.ResolveNameXText(n.SheetRefIndex, n.NameIndex);
        }

        public String GetSheetFirstNameByExternSheet(int externSheetIndex)
        {
            return _iBook.FindSheetFirstNameFromExternSheet(externSheetIndex);
        }
        public String GetSheetLastNameByExternSheet(int externSheetIndex)
        {
            return _iBook.FindSheetLastNameFromExternSheet(externSheetIndex);
        }
        public String GetNameText(NamePtg namePtg)
        {
            return _iBook.GetNameRecord(namePtg.Index).NameText;
        }
        public IEvaluationName GetName(NamePtg namePtg)
        {
            int ix = namePtg.Index;
            return new Name(_iBook.GetNameRecord(ix), ix);
        }
        public Ptg[] GetFormulaTokens(IEvaluationCell evalCell)
        {
            ICell cell = ((HSSFEvaluationCell)evalCell).HSSFCell;
            //if (false)
            //{
            //    // re-parsing the formula text also works, but is a waste of time
            //    // It is useful from time to time to run all unit tests with this code
            //    // to make sure that all formulas POI can evaluate can also be parsed.
            //    try
            //    {
            //        return HSSFFormulaParser.Parse(cell.CellFormula, _uBook, FormulaType.Cell, _uBook.GetSheetIndex(cell.GetSheet()));
            //    }
            //    catch (FormulaParseException e)
            //    {
            //        // Note - as of Bugzilla 48036 (svn r828244, r828247) POI is capable of evaluating
            //        // IntesectionPtg.  However it is still not capable of parsing it.
            //        // So FormulaEvalTestData.xls now contains a few formulas that produce errors here.
            //        logger.Log(POILogger.ERROR, e.Message);
            //    }
            //}
            FormulaRecordAggregate fr = (FormulaRecordAggregate)((HSSFCell)cell).CellValueRecord;
            return fr.FormulaTokens;
        }

        public UDFFinder GetUDFFinder()
        {
            return _uBook.GetUDFFinder();
        }


        private class Name : IEvaluationName
        {

            private NameRecord _nameRecord;
            private int _index;

            public Name(NameRecord nameRecord, int index)
            {
                _nameRecord = nameRecord;
                _index = index;
            }
            public Ptg[] NameDefinition
            {
                get{
                    return _nameRecord.NameDefinition;
                }
            }
            public String NameText
            {
                get{
                    return _nameRecord.NameText;
                }
            }
            public bool HasFormula
            {
                get{
                    return _nameRecord.HasFormula;
                }
            }
            public bool IsFunctionName
            {
                get{
                    return _nameRecord.IsFunctionName;
                }
            }
            public bool IsRange
            {
                get
                {
                    return _nameRecord.HasFormula; // TODO - is this right?
                }
            }
            public NamePtg CreatePtg()
            {
                return new NamePtg(_index);
            }
        }
        private int GetSheetExtIx(SheetIdentifier sheetIden)
        {
            int extIx;
            if (sheetIden == null)
            {
                extIx = -1;
            }
            else
            {
                String workbookName = sheetIden.BookName;
                String firstSheetName = sheetIden.SheetId.Name;
                String lastSheetName = firstSheetName;

                if (sheetIden is SheetRangeIdentifier)
                {
                    lastSheetName = ((SheetRangeIdentifier)sheetIden).LastSheetIdentifier.Name;
                }

                if (workbookName == null)
                {
                    int firstSheetIndex = _uBook.GetSheetIndex(firstSheetName);
                    int lastSheetIndex = _uBook.GetSheetIndex(lastSheetName);
                    extIx = _iBook.checkExternSheet(firstSheetIndex, lastSheetIndex);
                }
                else
                {
                    extIx = _iBook.GetExternalSheetIndex(workbookName, firstSheetName, lastSheetName);
                }
            }
            return extIx;
        }
        public SpreadsheetVersion GetSpreadsheetVersion()
        {
            return SpreadsheetVersion.EXCEL97;
        }
    }
}