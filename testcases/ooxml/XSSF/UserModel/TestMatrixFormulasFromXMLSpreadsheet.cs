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

namespace TestCases.XSSF.UserModel
{
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using NUnit.Framework.Legacy;
    using System.Globalization;
    using TestCases.SS.Formula.Functions;
    [TestFixture]
    public sealed class TestMatrixFormulasFromXMLSpreadsheet
    {
        private static POILogger LOG = POILogFactory.GetLogger(typeof(TestMatrixFormulasFromXMLSpreadsheet));
        
        private static XSSFWorkbook workbook;
        private static ISheet sheet;
        private static IFormulaEvaluator evaluator;
        private static CultureInfo userLocale;

        /*
         * Unlike TestFormulaFromSpreadsheet which this class is modified from, there is no
         * differentiation between operators and functions, if more functionality is implemented with
         * array formulas then it might be worth it to separate operators from functions
         * 
         * Also, output matrices are statically 3x3, if larger matrices wanted to be tested
         * then adding matrix size parameter would be useful and parsing would be based off that.
         */

        private static class Navigator
        {
            /// <summary>
            /// Name of the test spreadsheet (found in the standard test data folder)
            /// </summary>
            public static readonly string FILENAME = "MatrixFormulaEvalTestData.xlsx";
            /// <summary>
            /// Row (zero-based) in the spreadsheet where operations start
            /// </summary>
            public static readonly int START_OPERATORS_ROW_INDEX = 1;
            /// <summary>
            /// Column (zero-based) in the spreadsheet where operations start
            /// </summary>
            public static readonly int START_OPERATORS_COL_INDEX = 0;
            /// <summary>
            /// Column (zero-based) in the spreadsheet where evaluations start
            /// </summary>
            public static readonly int START_RESULT_COL_INDEX = 7;
            /// <summary>
            /// Column separation in the spreadsheet between evaluations and expected results
            /// </summary>
            public static readonly int COL_OFF_EXPECTED_RESULT = 3;
            /// <summary>
            /// Row separation in the spreadsheet between operations
            /// </summary>
            public static readonly int ROW_OFF_NEXT_OP = 4;
            /// <summary>
            /// Used to indicate when there are no more operations left
            /// </summary>
            public static readonly string END_OF_TESTS = "<END>";

        }
        [OneTimeTearDown]
        public static void closeResource()
        {
            LocaleUtil.SetUserLocale(userLocale);
            workbook.Close();
        }


        public static List<object[]> data()
        {
            // Function "Text" uses custom-formats which are locale specific
            // can't Set the locale on a per-testrun execution, as some Settings have been
            // already Set, when we would try to change the locale by then
            userLocale = LocaleUtil.GetUserLocale();
            LocaleUtil.SetUserLocale(CultureInfo.CurrentCulture);

            workbook = XSSFTestDataSamples.OpenSampleWorkbook(Navigator.FILENAME);
            sheet = workbook.GetSheetAt(0);
            evaluator = new XSSFFormulaEvaluator(workbook);

            List<object[]> data = new List<object[]>();

            processFunctionGroup(data, Navigator.START_OPERATORS_ROW_INDEX, null);

            return data;
        }

        /// <summary>
        /// </summary>
        /// <param name="startRowIndex">row index in the spreadsheet where the first function/operator is found</param>
        /// <param name="testFocusFunctionName">name of a single function/operator to test alone.
        /// Typically pass <c>null</c> to test all functions
        /// </param>
        private static void processFunctionGroup(List<object[]> data, int startRowIndex, string testFocusFunctionName)
        {
            for(int rowIndex = startRowIndex; true; rowIndex += Navigator.ROW_OFF_NEXT_OP)
            {
                IRow r = sheet.GetRow(rowIndex);
                string targetFunctionName = GetTargetFunctionName(r);
                ClassicAssert.IsNotNull("Test spreadsheet cell empty on row ("
                        + (rowIndex) + "). Expected function name or '"
                        + Navigator.END_OF_TESTS + "'", targetFunctionName);
                if(targetFunctionName.Equals(Navigator.END_OF_TESTS))
                {
                    // found end of functions list
                    break;
                }
                if(testFocusFunctionName == null || targetFunctionName.Equals(testFocusFunctionName, StringComparison.OrdinalIgnoreCase))
                {
                    data.Add(new object[] { targetFunctionName, rowIndex });
                }
            }
        }

        [TestCaseSource(nameof(data))]
        public void ProcessFunctionRow(String targetFunctionName, int formulasRowIdx)
        {

            int endColNum = Navigator.START_RESULT_COL_INDEX + Navigator.COL_OFF_EXPECTED_RESULT;

            for(int rowNum = formulasRowIdx; rowNum < formulasRowIdx + Navigator.ROW_OFF_NEXT_OP - 1; rowNum++)
            {
                for(int colNum = Navigator.START_RESULT_COL_INDEX; colNum < endColNum; colNum++)
                {
                    IRow r = sheet.GetRow(rowNum);

                    /* mainly to escape row Assert.Failures on MDETERM which only returns a scalar */
                    if(r == null)
                    {
                        continue;
                    }

                    ICell c = sheet.GetRow(rowNum).GetCell(colNum);

                    if(c == null || c.CellType != CellType.Formula)
                    {
                        continue;
                    }

                    CellValue actValue = evaluator.Evaluate(c);
                    ICell expValue = sheet.GetRow(rowNum).GetCell(colNum + Navigator.COL_OFF_EXPECTED_RESULT);

                    string msg = String.Format("Function '{0}': Formula: {1} @ {2}:{3}"
                       , targetFunctionName, c.CellFormula, rowNum, colNum);

                    ClassicAssert.IsNotNull(expValue, msg + " - Bad Setup data expected value is null");
                    ClassicAssert.IsNotNull(actValue, msg + " - actual value was null");

                    CellType cellType = expValue.CellType;
                    switch(cellType)
                    {
                        case CellType.Blank:
                            ClassicAssert.AreEqual(CellType.Blank, actValue.CellType, msg);
                            break;
                        case CellType.Boolean:
                            ClassicAssert.AreEqual(CellType.Boolean, actValue.CellType, msg);
                            ClassicAssert.AreEqual(expValue.BooleanCellValue, actValue.BooleanValue, msg);
                            break;
                        case CellType.Error:
                            ClassicAssert.AreEqual(CellType.Error, actValue.CellType, msg);
                            ClassicAssert.AreEqual(ErrorEval.GetText(expValue.ErrorCellValue), ErrorEval.GetText(actValue.ErrorValue), msg);
                            break;
                        case CellType.Formula: // will never be used, since we will call method After formula evaluation
                            ClassicAssert.Fail("Cannot expect formula as result of formula evaluation: " + msg);
                            break;
                        case CellType.Numeric:
                            ClassicAssert.AreEqual(CellType.Numeric, actValue.CellType, msg);
                            TestMathX.AssertEquals(msg, expValue.NumericCellValue, actValue.NumberValue, TestMathX.POS_ZERO, TestMathX.DIFF_TOLERANCE_FACTOR);
                            break;
                        case CellType.String:
                            ClassicAssert.AreEqual(CellType.String, actValue.CellType, msg);
                            ClassicAssert.AreEqual(expValue.RichStringCellValue.String, actValue.StringValue, msg);
                            break;
                        default:
                            ClassicAssert.Fail("Unexpected cell type: " + cellType);
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <returns><c>null</c> if cell is missing, empty or blank</returns>
        private static string GetTargetFunctionName(IRow r)
        {
            if(r == null)
            {
                LOG.Log(POILogger.WARN, "Warning - given null row, can't figure out function name");
                return null;
            }
            ICell cell = r.GetCell(Navigator.START_OPERATORS_COL_INDEX);
            LOG.Log(POILogger.DEBUG, Navigator.START_OPERATORS_COL_INDEX);
            if(cell == null)
            {
                LOG.Log(POILogger.WARN, "Warning - Row " + r.RowNum + " has no cell " + Navigator.START_OPERATORS_COL_INDEX + ", can't figure out function name");
                return null;
            }
            if(cell.CellType == CellType.Blank)
            {
                return null;
            }
            if(cell.CellType == CellType.String)
            {
                return cell.RichStringCellValue.String;
            }

            throw new AssertFailedException("Bad cell type for 'function name' column: ("
                    + cell.CellType + ") row (" + (r.RowNum +1) + ")");
        }
    }
}


