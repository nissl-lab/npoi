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
using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFEvaluationWorkbook : BaseXSSFEvaluationWorkbook
    {
        private readonly SXSSFWorkbook _xBook;

        public static SXSSFEvaluationWorkbook Create(SXSSFWorkbook book)
        {
            if (book == null)
            {
                return null;
            }
            return new SXSSFEvaluationWorkbook(book);
        }

        private SXSSFEvaluationWorkbook(SXSSFWorkbook book) : base(book.XssfWorkbook)
        {
            _xBook = book;
        }


        public override int GetSheetIndex(IEvaluationSheet evalSheet)
        {
            SXSSFSheet sheet = ((SXSSFEvaluationSheet)evalSheet).GetSXSSFSheet();
            return _xBook.GetSheetIndex(sheet);
        }


        public override IEvaluationSheet GetSheet(int sheetIndex)
        {
            return new SXSSFEvaluationSheet(_xBook.GetSheetAt(sheetIndex) as SXSSFSheet);
        }


        public override Ptg[] GetFormulaTokens(IEvaluationCell evalCell)
        {
            SXSSFCell cell = ((SXSSFEvaluationCell)evalCell).GetSXSSFCell();
            return FormulaParser.Parse(cell.CellFormula, this, FormulaType.Cell, _xBook.GetSheetIndex(cell.Sheet));
        }
    }

}
