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
    using NPOI.XSSF.Model;
    using NPOI.SS.Formula;
    using System;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.PTG;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS;
    using NPOI.SS.Formula.Udf;


    /**
     * Internal POI use only
     *
     * @author Josh Micich
     */
    public class XSSFEvaluationWorkbook : IFormulaRenderingWorkbook, IEvaluationWorkbook, FormulaParsingWorkbook
    {

        private XSSFWorkbook _uBook;

        public static XSSFEvaluationWorkbook Create(XSSFWorkbook book)
        {
            if (book == null)
            {
                return null;
            }
            return new XSSFEvaluationWorkbook(book);
        }

        private XSSFEvaluationWorkbook(XSSFWorkbook book)
        {
            _uBook = book;
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

        public IEvaluationName GetName(String name, int sheetIndex)
        {
            for (int i = 0; i < _uBook.GetNumberOfNames(); i++)
            {
                XSSFName nm = _uBook.GetNameAt(i);
                String nameText = nm.GetNameName();
                if (name.Equals(nameText, StringComparison.InvariantCultureIgnoreCase) && nm.GetSheetIndex() == sheetIndex)
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
            throw new NotImplementedException();
        }

        public NameXPtg GetNameXPtg(String name)
        {
            IndexedUDFFinder udfFinder = (IndexedUDFFinder)GetUDFFinder();
            FreeRefFunction func = udfFinder.FindFunction(name);
            if (func == null) return null;
            else return new NameXPtg(0, udfFinder.GetFunctionIndex(name));
        }

        public String ResolveNameXText(NameXPtg n)
        {
            int idx = n.NameIndex;
            IndexedUDFFinder udfFinder = (IndexedUDFFinder)GetUDFFinder();
            return udfFinder.GetFunctionName(idx);
        }

        public IEvaluationSheet GetSheet(int sheetIndex)
        {
            return new XSSFEvaluationSheet(_uBook.GetSheetAt(sheetIndex));
        }

        public ExternalSheet GetExternalSheet(int externSheetIndex)
        {
            // TODO Auto-generated method stub
            return null;
        }
        public int GetExternalSheetIndex(String workbookName, String sheetName)
        {
            throw new Exception("not implemented yet");
        }
        public int GetSheetIndex(String sheetName)
        {
            return _uBook.GetSheetIndex(sheetName);
        }

        public String GetSheetNameByExternSheet(int externSheetIndex)
        {
            int sheetIndex = ConvertFromExternalSheetIndex(externSheetIndex);
            return _uBook.GetSheetName(sheetIndex);
        }

        public String GetNameText(NamePtg namePtg)
        {
            return _uBook.GetNameAt(namePtg.Index).GetNameName();
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
            String formulaText = CleanXSSFFormulaText(cell.GetCellFormula());
            return FormulaParser.Parse(formulaText, frBook, FormulaType.CELL, _uBook.GetSheetIndex(cell.Sheet));
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
            private FormulaParsingWorkbook _fpBook;

            public Name(XSSFName name, int index, FormulaParsingWorkbook fpBook)
            {
                _nameRecord = name;
                _index = index;
                _fpBook = fpBook;
            }

            public Ptg[] GetNameDefInition()
            {

                return FormulaParser.Parse(_nameRecord.GetRefersToFormula(), _fpBook, FormulaType.NAMEDRANGE, _nameRecord.GetSheetIndex());
            }

            public String GetNameText()
            {
                return _nameRecord.GetNameName();
            }

            public bool HasFormula()
            {
                // TODO - no idea if this is right
                CT_DefinedName ctn = _nameRecord.GetCTName();
                String strVal = ctn.Value;
                return !ctn.function && strVal != null && strVal.Length > 0;
            }

            public bool IsFunctionName()
            {
                return _nameRecord.IsFunctionName();
            }

            public bool IsRange()
            {
                return HasFormula(); // TODO - is this right?
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

