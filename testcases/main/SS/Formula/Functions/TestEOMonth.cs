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
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    [TestFixture]
    public class TestEOMonth
    {

        private static double BAD_DATE = -1.0;

        private static double DATE_1900_01_01 = 1.0;
        private static double DATE_1900_01_31 = 31.0;
        private static double DATE_1900_02_28 = 59.0;
        private static double DATE_1902_09_26 = 1000.0;
        private static double DATE_1902_09_30 = 1004.0;
        private static double DATE_2034_06_09 = 49104.0;
        private static double DATE_2034_06_30 = 49125.0;
        private static double DATE_2034_07_31 = 49156.0;

        private FreeRefFunction eOMonth = EOMonth.instance;
        private OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);

        [Test]
        public void TestEOMonthProperValues()
        {
            // verify some border-case combinations of startDate and month-increase
            CheckValue(DATE_1900_01_01, 0, DATE_1900_01_31);
            CheckValue(DATE_1900_01_01, 1, DATE_1900_02_28);
            CheckValue(DATE_1902_09_26, 0, DATE_1902_09_30);
            CheckValue(DATE_2034_06_09, 0, DATE_2034_06_30);
            CheckValue(DATE_2034_06_09, 1, DATE_2034_07_31);
        }

        [Test]
        public void TestEOMonthBadDateValues()
        {
            CheckValue(0.0, -2, BAD_DATE);
            CheckValue(0.0, -3, BAD_DATE);
            CheckValue(DATE_1900_01_31, -1, BAD_DATE);
        }

        private void CheckValue(double startDate, int monthInc, double expectedResult)
        {
            NumberEval result = (NumberEval)eOMonth.Evaluate(new ValueEval[] { new NumberEval(startDate), new NumberEval(monthInc) }, ec);
            Assert.AreEqual(expectedResult, result.NumberValue);
        }

        [Test]
        public void TestEOMonthZeroDate()
        {
            NumberEval result = (NumberEval)eOMonth.Evaluate(new ValueEval[] { new NumberEval(0), new NumberEval(0) }, ec);
            Assert.AreEqual(DATE_1900_01_31, result.NumberValue, "0 startDate is 1900-01-00");

            result = (NumberEval)eOMonth.Evaluate(new ValueEval[] { new NumberEval(0), new NumberEval(1) }, ec);
            Assert.AreEqual(DATE_1900_02_28, result.NumberValue, "0 startDate is 1900-01-00");
        }

        [Test]
        public void TestEOMonthInvalidArguments()
        {
            ValueEval result = eOMonth.Evaluate(new ValueEval[] { new NumberEval(DATE_1902_09_26) }, ec);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorConstants.ERROR_VALUE, ((ErrorEval)result).ErrorCode);

            result = eOMonth.Evaluate(new ValueEval[] { new StringEval("a"), new StringEval("b") }, ec);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorConstants.ERROR_VALUE, ((ErrorEval)result).ErrorCode);
        }

        [Test]
        public void TestEOMonthIncrease()
        {
            CheckOffset(DateTime.Now, 2);
        }

        [Test]
        public void TestEOMonthDecrease()
        {
            CheckOffset(DateTime.Now, -2);
        }

        private void CheckOffset(DateTime startDate, int offset)
        {
            NumberEval result = (NumberEval)eOMonth.Evaluate(new ValueEval[] { new NumberEval(DateUtil.GetExcelDate(startDate)), new NumberEval(offset) }, ec);
            DateTime resultDate = DateUtil.GetJavaDate(result.NumberValue);

            //the month
            DateTime dtEnd = startDate.AddMonths(offset);
            //next month
            dtEnd = dtEnd.AddMonths(1);
            //first day of next month
            dtEnd = new DateTime(dtEnd.Year, dtEnd.Month, 1);
            //last day of the month
            dtEnd = dtEnd.AddDays(-1);
            Assert.AreEqual(dtEnd, resultDate);
        }

        [Test]
        public void TestBug56688()
        {
            NumberEval result = (NumberEval)eOMonth.Evaluate(new ValueEval[] { new NumberEval(DATE_1902_09_26), new RefEvalImplementation(new NumberEval(0)) }, ec);
            Assert.AreEqual(DATE_1902_09_30, result.NumberValue);
        }

        [Test]
        public void TestRefEvalStartDate()
        {
            NumberEval result = (NumberEval)eOMonth.Evaluate(new ValueEval[] { new RefEvalImplementation(new NumberEval(DATE_1902_09_26)), new NumberEval(0) }, ec);
            Assert.AreEqual(DATE_1902_09_30, result.NumberValue);
        }

        [Test]
        public void TestEOMonthBlankValueEval()
        {
            NumberEval Evaluate = (NumberEval)eOMonth.Evaluate(new ValueEval[] { BlankEval.instance, new NumberEval(0) }, ec);
            Assert.AreEqual(DATE_1900_01_31, Evaluate.NumberValue, "Blank is handled as 0");
        }

        [Test]
        public void TestEOMonthBlankRefValueEval()
        {
            NumberEval result = (NumberEval)eOMonth.Evaluate(new ValueEval[] { new RefEvalImplementation(BlankEval.instance), new NumberEval(1) }, ec);
            Assert.AreEqual(DATE_1900_02_28, result.NumberValue, "Blank is handled as 0");

            result = (NumberEval)eOMonth.Evaluate(new ValueEval[] { new NumberEval(1), new RefEvalImplementation(BlankEval.instance) }, ec);
            Assert.AreEqual(DATE_1900_01_31, result.NumberValue, "Blank is handled as 0");
        }
    }

}