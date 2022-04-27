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
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.UserModel;
    using NUnit.Framework;

    /**
     * Tests for OFFSET function implementation
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestOffset
    {

        private static void ConfirmDoubleConvert(double doubleVal, int expected)
        {
            try
            {
                Assert.AreEqual(expected, Offset.EvaluateIntArg(new NumberEval(doubleVal), -1, -1));
            }
            catch (EvaluationException e)
            {
                throw new AssertionException("Unexpected error '" + e.GetErrorEval().ToString() + "'.");
            }
        }
        /**
         * Excel's double to int conversion (for function 'OFFSET()') behaves more like Math.floor().
         * Note - negative values are not symmetrical
         * Fractional values are silently tRuncated.
         * TRuncation is toward negative infInity.
         */
        [Test]
        public void TestDoubleConversion()
        {

            ConfirmDoubleConvert(100.09, 100);
            ConfirmDoubleConvert(100.01, 100);
            ConfirmDoubleConvert(100.00, 100);
            ConfirmDoubleConvert(99.99, 99);

            ConfirmDoubleConvert(+2.01, +2);
            ConfirmDoubleConvert(+2.00, +2);
            ConfirmDoubleConvert(+1.99, +1);
            ConfirmDoubleConvert(+1.01, +1);
            ConfirmDoubleConvert(+1.00, +1);
            ConfirmDoubleConvert(+0.99, 0);
            ConfirmDoubleConvert(+0.01, 0);
            ConfirmDoubleConvert(0.00, 0);
            ConfirmDoubleConvert(-0.01, -1);
            ConfirmDoubleConvert(-0.99, -1);
            ConfirmDoubleConvert(-1.00, -1);
            ConfirmDoubleConvert(-1.01, -2);
            ConfirmDoubleConvert(-1.99, -2);
            ConfirmDoubleConvert(-2.00, -2);
            ConfirmDoubleConvert(-2.01, -3);
        }
        [Test]
        public void TestLinearOffsetRange()
        {
            Offset.LinearOffsetRange lor;

            lor = new Offset.LinearOffsetRange(3, 2);
            Assert.AreEqual(3, lor.FirstIndex);
            Assert.AreEqual(4, lor.LastIndex);
            lor = lor.NormaliseAndTranslate(0); // expected no change
            Assert.AreEqual(3, lor.FirstIndex);
            Assert.AreEqual(4, lor.LastIndex);

            lor = lor.NormaliseAndTranslate(5);
            Assert.AreEqual(8, lor.FirstIndex);
            Assert.AreEqual(9, lor.LastIndex);

            // negative length

            lor = new Offset.LinearOffsetRange(6, -4).NormaliseAndTranslate(0);
            Assert.AreEqual(3, lor.FirstIndex);
            Assert.AreEqual(6, lor.LastIndex);


            // bounds Checking
            lor = new Offset.LinearOffsetRange(0, 100);
            Assert.IsFalse(lor.IsOutOfBounds(0, 16383));
            lor = lor.NormaliseAndTranslate(16300);
            Assert.IsTrue(lor.IsOutOfBounds(0, 16383));
            Assert.IsFalse(lor.IsOutOfBounds(0, 65535));
        }

        [Test]
        public void TestOffsetWithEmpty23Arguments()
        {
            IWorkbook workbook = new HSSFWorkbook();
            ICell cell = workbook.CreateSheet().CreateRow(0).CreateCell(0);
            cell.SetCellFormula("OFFSET(B1,,)");
            string value = "EXPECTED_VALUE";
            ICell valueCell = cell.Row.CreateCell(1);
            valueCell.SetCellValue(value);
            workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
            Assert.AreEqual(CellType.String, cell.CachedFormulaResultType);
            Assert.AreEqual(value, cell.StringCellValue);
        }
    }

}