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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.HSSF.UserModel;

    using TestCases.HSSF;
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using NPOI.SS.UserModel;

    /**
     * @author aviks
     * 
     * This Testcase contains Tests for bugs that are yet to be fixed. Therefore,
     * the standard ant Test target does not run these Tests. Run this Testcase with
     * the single-Test target. The names of the Tests usually correspond to the
     * Bugzilla id's PLEASE MOVE Tests from this class to TestBugs once the bugs are
     * fixed, so that they are then run automatically.
     */
    [TestFixture]
    public class TestUnfixedBugs
    {
        //In POI bugzilla, this bug is taged as "RESOLVED WON'T FIX"
        [Test]
        [Ignore("WON'T FIX")]
        public void Test43493()
        {
            // Has crazy corrupt sub-records on
            // a EmbeddedObjectRefSubRecord
            try
            {
                HSSFTestDataSamples.OpenSampleWorkbook("43493.xls");
            }
            catch (RecordFormatException e)
            {
                if (e.InnerException.InnerException is IndexOutOfRangeException)
                {
                    throw new AssertionException("Identified bug 43493");
                }
                throw e;
            }
        }

        [Test]
        [Ignore("TestUnfixedBugs")] 
        public void Test49612()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("49612.xls");
            ISheet sh = wb.GetSheetAt(0);
            IRow row = sh.GetRow(0);
            ICell c1 = row.GetCell(2);
            ICell d1 = row.GetCell(3);
            ICell e1 = row.GetCell(2);

            Assert.AreEqual("SUM(BOB+JIM)", c1.CellFormula);

            // Problem 1: java.lang.ArrayIndexOutOfBoundsException in NPOI.HSSF.Model.LinkTable$ExternalBookBlock.GetNameText
            Assert.AreEqual("SUM('49612.xls'!BOB+'49612.xls'!JIM)", d1.CellFormula);

            //Problem 2
            //junit.framework.ComparisonFailure:
            //Expected :SUM('49612.xls'!BOB+'49612.xls'!JIM)
            //Actual   :SUM(BOB+JIM)
            Assert.AreEqual("SUM('49612.xls'!BOB+'49612.xls'!JIM)", e1.CellFormula);

            HSSFFormulaEvaluator eval = new HSSFFormulaEvaluator(wb);
            Assert.AreEqual(30.0, eval.Evaluate(c1).NumberValue, "Evaluating c1");

            //Problem 3:  java.lang.Exception: Unexpected arg eval type (NPOI.HSSF.Record.Formula.Eval.NameXEval)
            Assert.AreEqual(30, eval.Evaluate(d1).NumberValue, "Evaluating d1");

            Assert.AreEqual(30, eval.Evaluate(e1).NumberValue, "Evaluating e1");
        }
    }
}