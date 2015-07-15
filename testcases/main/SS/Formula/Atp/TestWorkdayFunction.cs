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
using System.Collections.Generic;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Atp;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NUnit.Framework;
using System.Globalization;
namespace TestCases.SS.Formula.Atp
{
    [TestFixture]
    public class TestWorkdayFunction
    {

        private static SimpleDateFormat formatter = new SimpleDateFormat("yyyy/MM/dd");
        private static int OCTOBER = 10;
        private static int NOVEMBER = 11;
        private static int DECEMBER = 12;
        private static int JANUARY = 1;
        private static int SEPTEMBER = 9;
        private static int APRIL = 4;
        private static int MAY = 5;
        private static String STARTING_DATE = formatter.Format(new DateTime(2008, OCTOBER, 1), CultureInfo.CurrentCulture );

        private static String FIRST_HOLIDAY = formatter.Format(new DateTime(2008, NOVEMBER, 26), CultureInfo.CurrentCulture);

        private static String SECOND_HOLIDAY = formatter.Format(new DateTime(2008, DECEMBER, 4), CultureInfo.CurrentCulture);

        private static String THIRD_HOLIDAY = formatter.Format(new DateTime(2009, JANUARY, 21), CultureInfo.CurrentCulture);

        private static String RETROATIVE_HOLIDAY = formatter.Format(new DateTime(2008, SEPTEMBER, 29), CultureInfo.CurrentCulture);

        private static OperationEvaluationContext EC = new OperationEvaluationContext(null, null, 1, 1, 1, null);
        [Test]
        public void TestFailWhenNoArguments()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, WorkdayFunction.instance.Evaluate(new ValueEval[0], null));
        }

        [Test]
        public void TestFailWhenLessThan2Arguments()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, WorkdayFunction.instance.Evaluate(new ValueEval[1], null));
        }
        [Test]
        public void TestFailWhenMoreThan3Arguments()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, WorkdayFunction.instance.Evaluate(new ValueEval[4], null));
        }
        [Test]
        public void TestFailWhenArgumentsAreNotDatesNorNumbers()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, WorkdayFunction.instance.Evaluate(
                    new ValueEval[] { new StringEval("Potato"), new StringEval("Cucumber") }, EC));
        }
        [Test]
        public void TestReturnWorkdays()
        {
            Assert.AreEqual(new DateTime(2009, APRIL, 30), DateUtil.GetJavaDate(((NumberEval)WorkdayFunction.instance.Evaluate(new ValueEval[]{
                new StringEval(STARTING_DATE), new NumberEval(151) }, EC)).NumberValue));
        }

        [Test]
        public void TestReturnWorkdaysSpanningAWeekendSubtractingDays()
        {
            String startDate = "2013/09/30";
            int days = -1;
            String expectedWorkDay = "2013/09/27";
            StringEval stringEval = new StringEval(startDate);
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(new ValueEval[]{
                stringEval, new NumberEval(days) }, EC)).NumberValue;
            Assert.AreEqual(expectedWorkDay, formatter.Format(DateUtil.GetJavaDate(numberValue)));
        }

        [Test]
        public void TestReturnWorkdaysSpanningAWeekendAddingDays()
        {
            String startDate = "2013/09/27";
            int days = 1;
            String expectedWorkDay = "2013/09/30";
            StringEval stringEval = new StringEval(startDate);
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(new ValueEval[]{
                stringEval, new NumberEval(days) }, EC)).NumberValue;
            Assert.AreEqual(expectedWorkDay, formatter.Format(DateUtil.GetJavaDate(numberValue)));
        }

        [Test]
        public void TestReturnWorkdaysWhenStartIsWeekendAddingDays()
        {
            String startDate = "2013/10/06";
            int days = 1;
            String expectedWorkDay = "2013/10/07";
            StringEval stringEval = new StringEval(startDate);
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(new ValueEval[]{
                stringEval, new NumberEval(days) }, EC)).NumberValue;
            Assert.AreEqual(expectedWorkDay, formatter.Format(DateUtil.GetJavaDate(numberValue)));
        }

        [Test]
        public void TestReturnWorkdaysWhenStartIsWeekendSubtractingDays()
        {
            String startDate = "2013/10/06";
            int days = -1;
            String expectedWorkDay = "2013/10/04";
            StringEval stringEval = new StringEval(startDate);
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(new ValueEval[]{
                stringEval, new NumberEval(days) }, EC)).NumberValue;
            Assert.AreEqual(expectedWorkDay, formatter.Format(DateUtil.GetJavaDate(numberValue)));
        }
        [Test]
        public void TestReturnWorkdaysWithDaysTruncated()
        {
            Assert.AreEqual(new DateTime(2009, APRIL, 30), DateUtil.GetJavaDate(((NumberEval)WorkdayFunction.instance.Evaluate(new ValueEval[]{
                new StringEval(STARTING_DATE), new NumberEval(151.99999) }, EC)).NumberValue));
        }
        [Test]
        public void TestReturnRetroativeWorkday()
        {
            Assert.AreEqual(new DateTime(2008, SEPTEMBER, 23), DateUtil.GetJavaDate(((NumberEval)WorkdayFunction.instance.Evaluate(new ValueEval[]{
                new StringEval(STARTING_DATE), new NumberEval(-5), new StringEval(RETROATIVE_HOLIDAY) }, EC))
                    .NumberValue));
        }
        [Test]
        public void TestReturnNetworkdaysWithManyHolidays()
        {
            Assert.AreEqual(new DateTime(2009, MAY, 5), DateUtil.GetJavaDate(((NumberEval)WorkdayFunction.instance.Evaluate(new ValueEval[]{
                new StringEval(STARTING_DATE), new NumberEval(151),
                new MockAreaEval(new string[]{FIRST_HOLIDAY, SECOND_HOLIDAY, THIRD_HOLIDAY}) }, EC)).NumberValue));
        }

        private class MockAreaEval : AreaEvalBase
        {

            private List<ValueEval> holidays;

            public MockAreaEval(String[] holidays)
                : this(0, 0, 0, holidays.Length - 1)
            {
                ;
                this.holidays = new List<ValueEval>();
                foreach (String holiday in holidays)
                {
                    this.holidays.Add(new StringEval(holiday));
                }
            }

            public MockAreaEval(int firstRow, int firstColumn, int lastRow, int lastColumn)
                : base(firstRow, firstColumn, lastRow, lastColumn)
            {

            }

            public override ValueEval GetRelativeValue(int sheetIndex, int relativeRowIndex, int relativeColumnIndex)
            {
                return this.holidays[(relativeColumnIndex)];
            }
            public override ValueEval GetRelativeValue(int relativeRowIndex, int relativeColumnIndex)
            {
                return GetRelativeValue(-1, relativeRowIndex, relativeColumnIndex);
            }
            public override AreaEval Offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx)
            {
                return null;
            }

            public override TwoDEval GetColumn(int columnIndex)
            {
                return null;
            }

            public override TwoDEval GetRow(int rowIndex)
            {
                return null;
            }

        }
    }

}