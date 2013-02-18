/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/*
 * TestFormulaRecordAggregate.java
 *
 * Created on March 21, 2003, 12:32 AM
 */

namespace TestCases.HSSF.Record.Aggregates
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Aggregates;
    using NUnit.Framework;
    using NPOI.Util;
    using NPOI.SS.Formula.PTG;
    using NPOI.HSSF.Model;
    using NPOI.SS.Formula;
    using NPOI.SS.Util;

    /**
     *
     * @author  avik
     */
    [TestFixture]
    public class TestFormulaRecordAggregate
    {
        [Test]
        public void TestBasic()
        {
            FormulaRecord f = new FormulaRecord();
            f.SetCachedResultTypeString();
            StringRecord s = new StringRecord();
            s.String = ("abc");
            FormulaRecordAggregate fagg = new FormulaRecordAggregate(f, s, SharedValueManager.CreateEmpty());
            Assert.AreEqual("abc", fagg.StringValue);
        }
        /**
 * Sometimes a {@link StringRecord} appears after a {@link FormulaRecord} even though the
 * formula has evaluated to a text value.  This might be more likely to occur when the formula
 * <i>can</i> evaluate to a text value.<br/>
 * Bug 46213 attachment 22874 has such an extra {@link StringRecord} at stream offset 0x5765.
 * This file seems to open in Excel (2007) with no trouble.  When it is re-saved, Excel omits
 * the extra record.  POI should do the same.
 */
        [Test]
        public void TestExtraStringRecord_bug46213()
        {
            FormulaRecord fr = new FormulaRecord();
            fr.Value = (2.0);
            StringRecord sr = new StringRecord();
            sr.String = ("NA");
            SharedValueManager svm = SharedValueManager.CreateEmpty();
            FormulaRecordAggregate fra;

            try
            {
                fra = new FormulaRecordAggregate(fr, sr, svm);
            }
            catch (RecordFormatException e)
            {
                if ("String record was  supplied but formula record flag is not  set".Equals(e.Message))
                {
                    throw new AssertionException("Identified bug 46213");
                }
                throw e;
            }
            TestCases.HSSF.UserModel.RecordInspector.RecordCollector rc = new TestCases.HSSF.UserModel.RecordInspector.RecordCollector();
            fra.VisitContainedRecords(rc);
            Record[] vraRecs = rc.Records;
            Assert.AreEqual(1, vraRecs.Length);
            Assert.AreEqual(fr, vraRecs[0]);
        }
        [Test]
        public void TestArrayFormulas()
        {
            int rownum = 4;
            int colnum = 4;

            FormulaRecord fr = new FormulaRecord();
            fr.Row=(rownum);
            fr.Column=((short)colnum);

            FormulaRecordAggregate agg = new FormulaRecordAggregate(fr, null, SharedValueManager.CreateEmpty());
            Ptg[] ptgsForCell = { new ExpPtg(rownum, colnum) };
            agg.SetParsedExpression(ptgsForCell);

            String formula = "SUM(A1:A3*B1:B3)";
            Ptg[] ptgs = HSSFFormulaParser.Parse(formula, null, FormulaType.Array, 0);
            agg.SetArrayFormula(new CellRangeAddress(rownum, rownum, colnum, colnum), ptgs);

            Assert.IsTrue(agg.IsPartOfArrayFormula);
            Assert.AreEqual("E5", agg.GetArrayFormulaRange().FormatAsString());
            Ptg[] ptg = agg.FormulaTokens;
            String fmlaSer = FormulaRenderer.ToFormulaString(null, ptg);
            Assert.AreEqual(formula, fmlaSer);

            agg.RemoveArrayFormula(rownum, colnum);
            Assert.IsFalse(agg.IsPartOfArrayFormula);
        }
    }
}