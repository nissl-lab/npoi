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

namespace TestCases.SS.Formula.Eval.Forked
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Eval.Forked;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using System;

    /**
     * @author Josh Micich
     */
    [TestFixture]
    public class TestForkedEvaluator
    {
        protected IWorkbook NewWorkbook()
        {
            return new HSSFWorkbook();
        }
        /**
         * Set up a calculation workbook with input cells nicely segregated on a
         * sheet called "Inputs"
         */
        private IWorkbook CreateWorkbook()
        {
            IWorkbook wb = NewWorkbook();
            ISheet sheet1 = wb.CreateSheet("Inputs");
            ISheet sheet2 = wb.CreateSheet("Calculations");
            IRow row;
            row = sheet2.CreateRow(0);
            row.CreateCell(0).CellFormula = ("B1*Inputs!A1-Inputs!B1");
            row.CreateCell(1).SetCellValue(5.0); // Calculations!B1

            // some default input values
            row = sheet1.CreateRow(0);
            row.CreateCell(0).SetCellValue(2.0); // Inputs!A1
            row.CreateCell(1).SetCellValue(3.0); // Inputs!B1
            return wb;
        }
        private class StabilityClassifier : IStabilityClassifier
        {
            public override bool IsCellFinal(int sheetIndex, int rowIndex, int columnIndex)
            {
                return sheetIndex == 1;
            }
        }
        /**
         * Shows a basic use-case for {@link ForkedEvaluator}
         */
        [Test]
        public void TestBasic()
        {
            IWorkbook wb = CreateWorkbook();

            // The stability classifier is useful to reduce memory consumption of caching logic
            IStabilityClassifier stabilityClassifier = new StabilityClassifier();

            ForkedEvaluator fe1 = ForkedEvaluator.Create(wb, stabilityClassifier, null);
            ForkedEvaluator fe2 = ForkedEvaluator.Create(wb, stabilityClassifier, null);

            // fe1 and fe2 can be used concurrently on separate threads

            fe1.UpdateCell("Inputs", 0, 0, new NumberEval(4.0));
            fe1.UpdateCell("Inputs", 0, 1, new NumberEval(1.1));

            fe2.UpdateCell("Inputs", 0, 0, new NumberEval(1.2));
            fe2.UpdateCell("Inputs", 0, 1, new NumberEval(2.0));

            Assert.AreEqual(18.9, ((NumberEval)fe1.Evaluate("Calculations", 0, 0)).NumberValue, 0.0);
            Assert.AreEqual(4.0, ((NumberEval)fe2.Evaluate("Calculations", 0, 0)).NumberValue, 0.0);
            fe1.UpdateCell("Inputs", 0, 0, new NumberEval(3.0));
            Assert.AreEqual(13.9, ((NumberEval)fe1.Evaluate("Calculations", 0, 0)).NumberValue, 0.0);

            wb.Close();
        }

        /**
         * As of Sep 2009, the Forked Evaluator can update values from existing cells (this is because
         * the underlying 'master' cell is used as a key into the calculation cache.  Prior to the fix
         * for this bug, an attempt to update a missing cell would result in NPE.  This junit Tests for
         * a more meaningful error message.<br/>
         *
         * An alternate solution might involve allowing empty cells to be Created as necessary.  That
         * was considered less desirable because so far, the underlying 'master' workbook is strictly
         * <i>read-only</i> with respect to the ForkedEvaluator.
         */
        [Test]
        public void TestMissingInputCellH()
        {
            IWorkbook wb = CreateWorkbook();
            ForkedEvaluator fe = ForkedEvaluator.Create(wb, null, null);
            // attempt update input at cell A2 (which is missing)
            try
            {
                fe.UpdateCell("Inputs", 1, 0, new NumberEval(4.0));
                throw new AssertionException(
                        "Expected exception to be thrown due to missing input cell");
            }
            catch (NullReferenceException e)
            {
                if (e.TargetSite.Equals("IdentityKey"))
                {
                    throw new AssertionException("Identified bug with update of missing input cell");
                }
                throw e;
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.Equals(
                        "Underlying cell 'A2' is missing in master sheet."))
                {
                    // expected during successful Test
                }
                else
                {
                    throw e;
                }
            }
        }
    }

}