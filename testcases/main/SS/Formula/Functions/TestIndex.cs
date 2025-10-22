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

    using System;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Util;
    using NPOI.Util;
    using TestCases.SS.Formula.Functions;
    using NPOI.SS.Formula.Functions;
    using Index = NPOI.SS.Formula.Functions.Index;
    using TestCases.HSSF;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /**
     * Tests for the INDEX() function.</p>
     *
     * This class Contains just a few specific cases that directly invoke {@link Index},
     * with minimum overhead.<br/>
     * Another Test: {@link TestIndexFunctionFromSpreadsheet} operates from a higher level
     * and has far greater coverage of input permutations.<br/>
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestIndex
    {

        private static Index FUNC_INST = new Index();
        private static double[] TEST_VALUES0 = {
            1, 2,
            3, 4,
            5, 6,
            7, 8,
            9, 10,
            11, 12,
    };

        /**
         * For the case when the first argument to INDEX() is an area reference
         */
        [Test]
        public void TestEvaluateAreaReference()
        {

            double[] values = TEST_VALUES0;
            ConfirmAreaEval("C1:D6", values, 4, 1, 7);
            ConfirmAreaEval("C1:D6", values, 6, 2, 12);
            ConfirmAreaEval("C1:D6", values, 3, 1, 5);

            // now treat same data as 3 columns, 4 rows
            ConfirmAreaEval("C10:E13", values, 2, 2, 5);
            ConfirmAreaEval("C10:E13", values, 4, 1, 10);
        }

        /**
         * @param areaRefString in Excel notation e.g. 'D2:E97'
         * @param dValues array of Evaluated values for the area reference
         * @param rowNum 1-based
         * @param colNum 1-based, pass -1 to signify argument not present
         */
        private static void ConfirmAreaEval(String areaRefString, double[] dValues,
                int rowNum, int colNum, double expectedResult)
        {
            ValueEval[] values = new ValueEval[dValues.Length];
            for (int i = 0; i < values.Length; i++)
            {
                values[i] = new NumberEval(dValues[i]);
            }
            AreaEval arg0 = EvalFactory.CreateAreaEval(areaRefString, values);

            ValueEval[] args;
            if (colNum > 0)
            {
                args = new ValueEval[] { arg0, new NumberEval(rowNum), new NumberEval(colNum), };
            }
            else
            {
                args = new ValueEval[] { arg0, new NumberEval(rowNum), };
            }

            double actual = invokeAndDereference(args);
            ClassicAssert.AreEqual(expectedResult, actual, 0D);
        }

        private static double invokeAndDereference(ValueEval[] args)
        {
            ValueEval ve = FUNC_INST.Evaluate(args, -1, -1);
            ve = WorkbookEvaluator.DereferenceResult(ve, -1, -1);
            ClassicAssert.AreEqual(typeof(NumberEval), ve.GetType());
            return ((NumberEval)ve).NumberValue;
        }

        /**
         * Tests expressions like "INDEX(A1:C1,,2)".<br/>
         * This problem was found while fixing bug 47048 and is observable up to svn r773441.
         */
        [Test]
        public void TestMissingArg()
        {
            ValueEval[] values = {
                new NumberEval(25.0),
                new NumberEval(26.0),
                new NumberEval(28.0),
        };
            AreaEval arg0 = EvalFactory.CreateAreaEval("A10:C10", values);
            ValueEval[] args = new ValueEval[] { arg0, MissingArgEval.instance, new NumberEval(2), };
            ValueEval actualResult;
            try
            {
                actualResult = FUNC_INST.Evaluate(args, -1, -1);
            }
            catch (Exception e)
            {
                if (e.Message.Equals("Unexpected arg eval type (NPOI.hssf.Record.Formula.Eval.MissingArgEval"))
                {
                    throw new AssertionException("Identified bug 47048b - INDEX() should support missing-arg");
                }
                throw;
            }
            // result should be an area eval "B10:B10"
            AreaEval ae = ConfirmAreaEval("B10:B10", actualResult);
            actualResult = ae.GetValue(0, 0);
            ClassicAssert.AreEqual(typeof(NumberEval), actualResult.GetType());
            ClassicAssert.AreEqual(26.0, ((NumberEval)actualResult).NumberValue, 0.0);
        }

        /**
         * When the argument to INDEX is a reference, the result should be a reference
         * A formula like "OFFSET(INDEX(A1:B2,2,1),1,1,1,1)" should return the value of cell B3.
         * This works because the INDEX() function returns a reference to A2 (not the value of A2)
         */
        [Test]
        public void TestReferenceResult()
        {
            ValueEval[] values = new ValueEval[4];
            Arrays.Fill(values, NumberEval.ZERO);
            AreaEval arg0 = EvalFactory.CreateAreaEval("A1:B2", values);
            ValueEval[] args = new ValueEval[] { arg0, new NumberEval(2), new NumberEval(1), };
            ValueEval ve = FUNC_INST.Evaluate(args, -1, -1);
            ConfirmAreaEval("A2:A2", ve);
        }

        /**
         * Confirms that the result is an area ref with the specified coordinates
         * @return <c>ve</c> cast to {@link AreaEval} if it is valid
         */
        private static AreaEval ConfirmAreaEval(String refText, ValueEval ve)
        {
            CellRangeAddress cra = CellRangeAddress.ValueOf(refText);
            ClassicAssert.IsTrue(ve is AreaEval);
            AreaEval ae = (AreaEval)ve;
            ClassicAssert.AreEqual(cra.FirstRow, ae.FirstRow);
            ClassicAssert.AreEqual(cra.FirstColumn, ae.FirstColumn);
            ClassicAssert.AreEqual(cra.LastRow, ae.LastRow);
            ClassicAssert.AreEqual(cra.LastColumn, ae.LastColumn);
            return ae;
        }

        [Test]
        [Ignore("bug 61116")]
        public void Test61859()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("maxindextest.xls");
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ISheet example1 = wb.GetSheetAt(0);
            ICell ex1cell1 = example1.GetRow(1).GetCell(6);
            ClassicAssert.AreEqual("MAX(INDEX(($B$2:$B$11=F2)*$A$2:$A$11,0))", ex1cell1.CellFormula);
            fe.Evaluate(ex1cell1);
            ClassicAssert.AreEqual(4.0, ex1cell1.NumericCellValue);

            ICell ex1cell2 = example1.GetRow(2).GetCell(6);
            ClassicAssert.AreEqual("MAX(INDEX(($B$2:$B$11=F3)*$A$2:$A$11,0))", ex1cell2.CellFormula);
            fe.Evaluate(ex1cell2);
            ClassicAssert.AreEqual(10.0, ex1cell2.NumericCellValue);

            ICell ex1cell3 = example1.GetRow(3).GetCell(6);
            ClassicAssert.AreEqual("MAX(INDEX(($B$2:$B$11=F4)*$A$2:$A$11,0))", ex1cell3.CellFormula);
            fe.Evaluate(ex1cell3);
            ClassicAssert.AreEqual(20.0, ex1cell3.NumericCellValue);
        }

        /// <summary>
        /// If both the Row_num and Column_num arguments are used,
        /// INDEX returns the value in the cell at the intersection of Row_num and Column_num
        /// </summary>
        [Test]
        public void TestReference2DArea()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            /// <summary>
            /// 1	2	3
            /// 4	5	6
            /// 7	8	9
            /// </summary>
            int val = 0;
            for(int i = 0; i < 3; i++)
            {
                IRow row = sheet.CreateRow(i);
                for(int j = 0; j < 3; j++)
                {
                    row.CreateCell(j).SetCellValue(++val);
                }
            }
            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ICell c1 = sheet.GetRow(0).CreateCell(5);
            c1.SetCellFormula("INDEX(A1:C3,2,2)");
            ICell c2 = sheet.GetRow(0).CreateCell(6);
            c2.SetCellFormula("INDEX(A1:C3,3,2)");

            ClassicAssert.AreEqual(5.0, fe.Evaluate(c1).NumberValue);
            ClassicAssert.AreEqual(8.0, fe.Evaluate(c2).NumberValue);
        }

        /// <summary>
        /// If Column_num is 0 (zero), INDEX returns the array of values for the entire row.
        /// </summary>
        [Test]
        public void TestArrayArgument_RowLookup()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            /// <summary>
            /// 1	2	3
            /// 4	5	6
            /// 7	8	9
            /// </summary>
            int val = 0;
            for(int i = 0; i < 3; i++)
            {
                IRow row = sheet.CreateRow(i);
                for(int j = 0; j < 3; j++)
                {
                    row.CreateCell(j).SetCellValue(++val);
                }
            }
            ICell c1 = sheet.GetRow(0).CreateCell(5);
            c1.SetCellFormula("SUM(INDEX(A1:C3,1,0))"); // sum of all values in the 1st row: 1 + 2 + 3 = 6

            ICell c2 = sheet.GetRow(0).CreateCell(6);
            c2.SetCellFormula("SUM(INDEX(A1:C3,2,0))"); // sum of all values in the 2nd row: 4 + 5 + 6 = 15

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ClassicAssert.AreEqual(6.0, fe.Evaluate(c1).NumberValue);
            ClassicAssert.AreEqual(15.0, fe.Evaluate(c2).NumberValue);

        }

        /// <summary>
        /// If Row_num is 0 (zero), INDEX returns the array of values for the entire column.
        /// </summary>
        [Test]
        public void TestArrayArgument_ColumnLookup()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            /// <summary>
            /// 1	2	3
            /// 4	5	6
            /// 7	8	9
            /// </summary>
            int val = 0;
            for(int i = 0; i < 3; i++)
            {
                IRow row = sheet.CreateRow(i);
                for(int j = 0; j < 3; j++)
                {
                    row.CreateCell(j).SetCellValue(++val);
                }
            }
            ICell c1 = sheet.GetRow(0).CreateCell(5);
            c1.SetCellFormula("SUM(INDEX(A1:C3,0,1))"); // sum of all values in the 1st column: 1 + 4 + 7 = 12

            ICell c2 = sheet.GetRow(0).CreateCell(6);
            c2.SetCellFormula("SUM(INDEX(A1:C3,0,3))"); // sum of all values in the 3rd column: 3 + 6 + 9 = 18

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ClassicAssert.AreEqual(12.0, fe.Evaluate(c1).NumberValue);
            ClassicAssert.AreEqual(18.0, fe.Evaluate(c2).NumberValue);
        }

        /// <summary>
        /// <para>
        /// =SUM(B1:INDEX(B1:B3,2))
        /// </para>
        /// <para>
        /// 	 The sum of the range starting at B1, and ending at the intersection of the 2nd row of the range B1:B3,
        /// 	 which is the sum of B1:B2.
        /// </para>
        /// </summary>
        [Test]
        public void TestDynamicReference()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            /// <summary>
            /// 1	2	3
            /// 4	5	6
            /// 7	8	9
            /// </summary>
            int val = 0;
            for(int i = 0; i < 3; i++)
            {
                IRow row = sheet.CreateRow(i);
                for(int j = 0; j < 3; j++)
                {
                    row.CreateCell(j).SetCellValue(++val);
                }
            }
            ICell c1 = sheet.GetRow(0).CreateCell(5);
            c1.SetCellFormula("SUM(B1:INDEX(B1:B3,2))"); // B1:INDEX(B1:B3,2) evaluates to B1:B2

            IFormulaEvaluator fe = wb.GetCreationHelper().CreateFormulaEvaluator();

            ClassicAssert.AreEqual(7.0, fe.Evaluate(c1).NumberValue);
        }
    }

}