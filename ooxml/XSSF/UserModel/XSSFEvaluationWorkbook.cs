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
using NPOI.SS.Formula.UDF;
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
    public class XSSFEvaluationWorkbook : BaseXSSFEvaluationWorkbook
    {
        private XSSFEvaluationSheet[] _sheetCache;

        public static XSSFEvaluationWorkbook Create(IWorkbook book)
        {
            if (book == null)
            {
                return null;
            }
            return new XSSFEvaluationWorkbook(book as XSSFWorkbook);
        }

        protected XSSFEvaluationWorkbook(XSSFWorkbook book)
            : base(book)
        {

        }

        public override void ClearAllCachedResultValues()
        {
            base.ClearAllCachedResultValues();
            _sheetCache = null;
        }

        public override int GetSheetIndex(IEvaluationSheet evalSheet)
        {
            XSSFSheet sheet = ((XSSFEvaluationSheet)evalSheet).XSSFSheet;
            return _uBook.GetSheetIndex(sheet);
        }

        public override IEvaluationSheet GetSheet(int sheetIndex)
        {
            // Performance optimization: build sheet cache the first time this is called
            // to avoid re-creating the XSSFEvaluationSheet each time a new cell is evaluated
            // EvaluationWorkbooks make not guarantee to synchronize changes made to
            // the underlying workbook after the EvaluationWorkbook is created.
            if (_sheetCache == null)
            {
                int numberOfSheets = _uBook.NumberOfSheets;
                _sheetCache = new XSSFEvaluationSheet[numberOfSheets];
                for (int i = 0; i < numberOfSheets; i++)
                {
                    _sheetCache[i] = new XSSFEvaluationSheet(_uBook.GetSheetAt(i));
                }
            }
            if (sheetIndex < 0 || sheetIndex >= _sheetCache.Length)
            {
                // do this to reuse the out-of-bounds logic and message from XSSFWorkbook
                _uBook.GetSheetAt(sheetIndex);
            }
            return _sheetCache[sheetIndex];
        }

        public override Ptg[] GetFormulaTokens(IEvaluationCell evalCell)
        {
            XSSFCell cell = ((XSSFEvaluationCell)evalCell).GetXSSFCell();
            int sheetIndex = _uBook.GetSheetIndex(cell.Sheet);
            int rowIndex = cell.RowIndex;
            return FormulaParser.Parse(cell.GetCellFormula(this), this, FormulaType.Cell, sheetIndex, rowIndex);
        }
    }
}

