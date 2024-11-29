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
    using NPOI.SS.Formula.Functions;
    using NUnit.Framework;
    using System;

    /**
     * Tests for ROW(), ROWS(), COLUMN(), COLUMNS()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRowCol
    {
        [Test]
        public void TestCol()
        {
            Function target = new Column();
            {
                ValueEval[] args = { EvalFactory.CreateRefEval("C5"), };
                double actual = NumericFunctionInvoker.Invoke(target, args);
                Assert.AreEqual(3, actual, 0D);
            }
            {
                ValueEval[] args = { EvalFactory.CreateAreaEval("E2:H12", new ValueEval[44]), };
                double actual = NumericFunctionInvoker.Invoke(target, args);
                Assert.AreEqual(5, actual, 0D);
            }
        }
        [Test]
        public void TestCell()
        {
            Function target = new Cell();
            {
                ValueEval[] args = { new StringEval("col"), EvalFactory.CreateRefEval("B5"), };
                double actual = NumericFunctionInvoker.Invoke(target, args);
                Assert.AreEqual(2, actual);
            }
            {
                ValueEval[] args = { new StringEval("row"), EvalFactory.CreateRefEval("B5"), };
                double actual = NumericFunctionInvoker.Invoke(target, args);
                Assert.AreEqual(5, actual);
            }
        }
        [Test]
        public void TestRow()
        {
            //throw new NotImplementedException();
            Function target = new Row();
            {
                ValueEval[] args = { EvalFactory.CreateRefEval("C5"), };
                double actual = NumericFunctionInvoker.Invoke(target, args);
                Assert.AreEqual(5, actual, 0D);
            }
            {
                ValueEval[] args = { EvalFactory.CreateAreaEval("E2:H12", new ValueEval[44]), };
                double actual = NumericFunctionInvoker.Invoke(target, args);
                Assert.AreEqual(2, actual, 0D);
            }
        }
        [Test]
        public void TestColumns()
        {

            ConfirmColumnsFunc("A1:F1", 6, 1);
            ConfirmColumnsFunc("A1:C2", 3, 2);
            ConfirmColumnsFunc("A1:B3", 2, 3);
            ConfirmColumnsFunc("A1:A6", 1, 6);

            ValueEval[] args = { EvalFactory.CreateRefEval("C5"), };
            double actual = NumericFunctionInvoker.Invoke(new Columns(), args);
            Assert.AreEqual(1, actual, 0D);
        }
        [Test]
        public void TestRows()
        {

            ConfirmRowsFunc("A1:F1", 6, 1);
            ConfirmRowsFunc("A1:C2", 3, 2);
            ConfirmRowsFunc("A1:B3", 2, 3);
            ConfirmRowsFunc("A1:A6", 1, 6);

            ValueEval[] args = { EvalFactory.CreateRefEval("C5"), };
            double actual = NumericFunctionInvoker.Invoke(new Rows(), args);
            Assert.AreEqual(1, actual, 0D);
        }

        private static void ConfirmRowsFunc(String areaRefStr, int nCols, int nRows)
        {
            ValueEval[] args = { EvalFactory.CreateAreaEval(areaRefStr, new ValueEval[nCols * nRows]), };

            double actual = NumericFunctionInvoker.Invoke(new Rows(), args);
            Assert.AreEqual(nRows, actual, 0D);
        }


        private static void ConfirmColumnsFunc(String areaRefStr, int nCols, int nRows)
        {
            ValueEval[] args = { EvalFactory.CreateAreaEval(areaRefStr, new ValueEval[nCols * nRows]), };

            double actual = NumericFunctionInvoker.Invoke(new Columns(), args);
            Assert.AreEqual(nCols, actual, 0D);
        }
    }

}