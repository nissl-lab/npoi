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

namespace TestCases.SS.Formula.Functions
{

    using NPOI.SS.Formula.Eval;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using System;
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;

    /**
     * Tests for the INDIRECT() function.</p>
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestIndirect
    {
        // convenient access to namespace
        //private static ErrorEval EE = null;

        private static void CreateDataRow(ISheet sheet, int rowIndex, params double[] vals)
        {
            IRow row = sheet.CreateRow(rowIndex);
            for (int i = 0; i < vals.Length; i++)
            {
                row.CreateCell(i).SetCellValue(vals[i]);
            }
        }

        private static HSSFWorkbook CreateWBA()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet1 = wb.CreateSheet("Sheet1");
            ISheet sheet2 = wb.CreateSheet("Sheet2");
            ISheet sheet3 = wb.CreateSheet("John's sales");

            CreateDataRow(sheet1, 0, 11, 12, 13, 14);
            CreateDataRow(sheet1, 1, 21, 22, 23, 24);
            CreateDataRow(sheet1, 2, 31, 32, 33, 34);

            CreateDataRow(sheet2, 0, 50, 55, 60, 65);
            CreateDataRow(sheet2, 1, 51, 56, 61, 66);
            CreateDataRow(sheet2, 2, 52, 57, 62, 67);

            CreateDataRow(sheet3, 0, 30, 31, 32);
            CreateDataRow(sheet3, 1, 33, 34, 35);

            IName name1 = wb.CreateName();
            name1.NameName = ("sales1");
            name1.RefersToFormula = ("Sheet1!A1:D1");

            IName name2 = wb.CreateName();
            name2.NameName = ("sales2");
            name2.RefersToFormula = ("Sheet2!B1:C3");

            IRow row = sheet1.CreateRow(3);
            row.CreateCell(0).SetCellValue("sales1");  //A4
            row.CreateCell(1).SetCellValue("sales2");  //B4

            return wb;
        }

        private static HSSFWorkbook CreateWBB()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet1 = wb.CreateSheet("Sheet1");
            ISheet sheet2 = wb.CreateSheet("Sheet2");
            ISheet sheet3 = wb.CreateSheet("## Look here!");

            CreateDataRow(sheet1, 0, 400, 440, 480, 520);
            CreateDataRow(sheet1, 1, 420, 460, 500, 540);

            CreateDataRow(sheet2, 0, 50, 55, 60, 65);
            CreateDataRow(sheet2, 1, 51, 56, 61, 66);

            CreateDataRow(sheet3, 0, 42);

            return wb;
        }
        [Test]
        public void TestBasic()
        {

            HSSFWorkbook wbA = CreateWBA();
            ICell c = wbA.GetSheetAt(0).CreateRow(5).CreateCell(2);
            HSSFFormulaEvaluator feA = new HSSFFormulaEvaluator(wbA);

            // non-error cases
            Confirm(feA, c, "INDIRECT(\"C2\")", 23);
            Confirm(feA, c, "INDIRECT(\"C2\", TRUE)", 23);
            Confirm(feA, c, "INDIRECT(\"$C2\")", 23);
            Confirm(feA, c, "INDIRECT(\"C$2\")", 23);
            Confirm(feA, c, "SUM(INDIRECT(\"Sheet2!B1:C3\"))", 351); // area ref
            Confirm(feA, c, "SUM(INDIRECT(\"Sheet2! B1 : C3 \"))", 351); // spaces in area ref
            Confirm(feA, c, "SUM(INDIRECT(\"'John''s sales'!A1:C1\"))", 93); // special chars in sheet name
            Confirm(feA, c, "INDIRECT(\"'Sheet1'!B3\")", 32); // redundant sheet name quotes
            Confirm(feA, c, "INDIRECT(\"sHeet1!B3\")", 32); // case-insensitive sheet name
            Confirm(feA, c, "INDIRECT(\" D3 \")", 34); // spaces around cell ref
            Confirm(feA, c, "INDIRECT(\"Sheet1! D3 \")", 34); // spaces around cell ref
            Confirm(feA, c, "INDIRECT(\"A1\", TRUE)", 11); // explicit arg1. only TRUE supported so far

            Confirm(feA, c, "INDIRECT(\"A1:G1\")", 13); // de-reference area ref (note formula is in C4)

            Confirm(feA, c, "SUM(INDIRECT(A4))", 50); // indirect defined name
            Confirm(feA, c, "SUM(INDIRECT(B4))", 351); // indirect defined name pointinh to other sheet

            // simple error propagation:

            // arg0 is Evaluated to text first
            Confirm(feA, c, "INDIRECT(#DIV/0!)", ErrorEval.DIV_ZERO);
            Confirm(feA, c, "INDIRECT(#DIV/0!)", ErrorEval.DIV_ZERO);
            Confirm(feA, c, "INDIRECT(#NAME?, \"x\")", ErrorEval.NAME_INVALID);
            Confirm(feA, c, "INDIRECT(#NUM!, #N/A)", ErrorEval.NUM_ERROR);

            // arg1 is Evaluated to bool before arg0 is decoded
            Confirm(feA, c, "INDIRECT(\"garbage\", #N/A)", ErrorEval.NA);
            Confirm(feA, c, "INDIRECT(\"garbage\", \"\")", ErrorEval.VALUE_INVALID); // empty string is not valid bool
            Confirm(feA, c, "INDIRECT(\"garbage\", \"flase\")", ErrorEval.VALUE_INVALID); // must be "TRUE" or "FALSE"


            // spaces around sheet name (with or without quotes Makes no difference)
            Confirm(feA, c, "INDIRECT(\"'Sheet1 '!D3\")", ErrorEval.REF_INVALID);
            Confirm(feA, c, "INDIRECT(\" Sheet1!D3\")", ErrorEval.REF_INVALID);
            Confirm(feA, c, "INDIRECT(\"'Sheet1' !D3\")", ErrorEval.REF_INVALID);


            Confirm(feA, c, "SUM(INDIRECT(\"'John's sales'!A1:C1\"))", ErrorEval.REF_INVALID); // bad quote escaping
            Confirm(feA, c, "INDIRECT(\"[Book1]Sheet1!A1\")", ErrorEval.REF_INVALID); // unknown external workbook
            Confirm(feA, c, "INDIRECT(\"Sheet3!A1\")", ErrorEval.REF_INVALID); // unknown sheet
#if !HIDE_UNREACHABLE_CODE
            if (false)
            { // TODO - support Evaluation of defined names
                Confirm(feA, c, "INDIRECT(\"Sheet1!IW1\")", ErrorEval.REF_INVALID); // bad column
                Confirm(feA, c, "INDIRECT(\"Sheet1!A65537\")", ErrorEval.REF_INVALID); // bad row
            }
#endif
            Confirm(feA, c, "INDIRECT(\"Sheet1!A 1\")", ErrorEval.REF_INVALID); // space in cell ref
        }
        [Test]
        public void TestMultipleWorkbooks()
        {
            HSSFWorkbook wbA = CreateWBA();
            ICell cellA = wbA.GetSheetAt(0).CreateRow(10).CreateCell(0);
            HSSFFormulaEvaluator feA = new HSSFFormulaEvaluator(wbA);

            HSSFWorkbook wbB = CreateWBB();
            ICell cellB = wbB.GetSheetAt(0).CreateRow(10).CreateCell(0);
            HSSFFormulaEvaluator feB = new HSSFFormulaEvaluator(wbB);

            String[] workbookNames = { "MyBook", "Figures for January", };
            HSSFFormulaEvaluator[] Evaluators = { feA, feB, };
            HSSFFormulaEvaluator.SetupEnvironment(workbookNames, Evaluators);

            Confirm(feB, cellB, "INDIRECT(\"'[Figures for January]## Look here!'!A1\")", 42); // same wb
            Confirm(feA, cellA, "INDIRECT(\"'[Figures for January]## Look here!'!A1\")", 42); // across workbooks

            // 2 level recursion
            Confirm(feB, cellB, "INDIRECT(\"[MyBook]Sheet2!A1\")", 50); // Set up (and Check) first level
            Confirm(feA, cellA, "INDIRECT(\"'[Figures for January]Sheet1'!A11\")", 50); // points to cellB
        }

        private static void Confirm(IFormulaEvaluator fe, ICell cell, String formula,
                double expectedResult)
        {
            fe.ClearAllCachedResultValues();
            cell.CellFormula = (formula);
            CellValue cv = fe.Evaluate(cell);
            if (cv.CellType != CellType.Numeric)
            {
                throw new AssertionException("expected numeric cell type but got " + cv.FormatAsString());
            }
            Assert.AreEqual(expectedResult, cv.NumberValue, 0.0);
        }
        private static void Confirm(IFormulaEvaluator fe, ICell cell, String formula,
                ErrorEval expectedResult)
        {
            fe.ClearAllCachedResultValues();
            cell.CellFormula=(formula);
            CellValue cv = fe.Evaluate(cell);
            if (cv.CellType != CellType.Error)
            {
                throw new AssertionException("expected error cell type but got " + cv.FormatAsString());
            }
            int expCode = expectedResult.ErrorCode;
            if (cv.ErrorValue != expCode)
            {
                throw new AssertionException("Expected error '" + ErrorEval.GetText(expCode)
                        + "' but got '" + cv.FormatAsString() + "'.");
            }
        }

        [Test]
        public void TestInvalidInput()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, Indirect.instance.Evaluate(new ValueEval[] { }, null));
        }
    }

}