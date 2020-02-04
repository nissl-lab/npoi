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
        private static String STARTING_DATE      = "2008/10/01";
        private static String FIRST_HOLIDAY      = "2008/11/26";
        private static String SECOND_HOLIDAY     = "2008/12/04";
        private static String THIRD_HOLIDAY      = "2009/01/21";
        private static String RETROATIVE_HOLIDAY = "2008/09/29";
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
            DateTime expDate = new DateTime(2009, 4, 30);
            ValueEval[] ve = { new StringEval(STARTING_DATE), new NumberEval(151) };
            DateTime actDate = DateUtil.GetJavaDate(((NumberEval)WorkdayFunction.instance.Evaluate(ve, EC)).NumberValue);
            Assert.AreEqual(expDate, actDate);
        }

        [Test]
        public void TestReturnWorkdaysSpanningAWeekendSubtractingDays()
        {
            //Calendar expCal = LocaleUtil.getLocaleCalendar(2013, 8, 27);
            //Date expDate = expCal.getTime();
            DateTime expDate = new DateTime(2013, 9, 27); //the month value of java date should plus one 

            ValueEval[] ve = { new StringEval("2013/09/30"), new NumberEval(-1) };
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(ve, EC)).NumberValue;
            Assert.AreEqual(41544.0, numberValue, 0);

            DateTime actDate = DateUtil.GetJavaDate(numberValue);
            Assert.AreEqual(expDate, actDate);
        }

        [Test]
        public void TestReturnWorkdaysSpanningAWeekendAddingDays()
        {
            //Calendar expCal = LocaleUtil.getLocaleCalendar(2013, 8, 30);
            //Date expDate = expCal.getTime();
            DateTime expDate = new DateTime(2013, 9, 30); //the month value of java date should plus one 

            ValueEval[] ve = { new StringEval("2013/09/27"), new NumberEval(1) };
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(ve, EC)).NumberValue;
            Assert.AreEqual(41547.0, numberValue, 0);

            DateTime actDate = DateUtil.GetJavaDate(numberValue);
            Assert.AreEqual(expDate, actDate);
        }

        [Test]
        public void TestReturnWorkdaysWhenStartIsWeekendAddingDays()
        {
            //Calendar expCal = LocaleUtil.getLocaleCalendar(2013, 9, 7);
            //Date expDate = expCal.getTime();
            DateTime expDate = new DateTime(2013, 10, 7); //the month value of java date should plus one 

            ValueEval[] ve = { new StringEval("2013/10/06"), new NumberEval(1) };
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(ve, EC)).NumberValue;
            Assert.AreEqual(41554.0, numberValue, 0);

            DateTime actDate = DateUtil.GetJavaDate(numberValue);
            Assert.AreEqual(expDate, actDate);

        }

        [Test]
        public void TestReturnWorkdaysWhenStartIsWeekendSubtractingDays()
        {
            //Calendar expCal = LocaleUtil.getLocaleCalendar(2013, 9, 4);
            //Date expDate = expCal.getTime();
            DateTime expDate = new DateTime(2013, 10, 4); //the month value of java date should plus one 

            ValueEval[] ve = { new StringEval("2013/10/06"), new NumberEval(-1) };
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(ve, EC)).NumberValue;
            Assert.AreEqual(41551.0, numberValue, 0);

            DateTime actDate = DateUtil.GetJavaDate(numberValue);
            Assert.AreEqual(expDate, actDate);
        }
        [Test]
        public void TestReturnWorkdaysWithDaysTruncated()
        {
            DateTime expDate = new DateTime(2009, 4, 30);
            ValueEval[] ve = { new StringEval(STARTING_DATE), new NumberEval(151.99999) };
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(ve, EC)).NumberValue;
            DateTime actDate = DateUtil.GetJavaDate(numberValue);
            Assert.AreEqual(expDate, actDate);
        }
        [Test]
        public void TestReturnRetroativeWorkday()
        {
            DateTime expDate = new DateTime(2008, 9, 23);
            ValueEval[] ve = new ValueEval[]{new StringEval(STARTING_DATE), new NumberEval(-5),
                new StringEval(RETROATIVE_HOLIDAY) };
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(ve, EC)).NumberValue;
            DateTime actDate = DateUtil.GetJavaDate(numberValue);
            Assert.AreEqual(expDate, actDate);
        }
        [Test]
        public void TestReturnNetworkdaysWithManyHolidays()
        {
            DateTime expDate = new DateTime(2009, 5, 5);
            ValueEval[] ve = {
                new StringEval(STARTING_DATE), new NumberEval(151),
                new MockAreaEval(new string[] {FIRST_HOLIDAY, SECOND_HOLIDAY, THIRD_HOLIDAY }) };
            double numberValue = ((NumberEval)WorkdayFunction.instance.Evaluate(ve, EC)).NumberValue;
            DateTime actDate = DateUtil.GetJavaDate(numberValue);
            Assert.AreEqual(expDate, actDate);
        }

        private class MockAreaEval : AreaEvalBase
        {

            private List<ValueEval> holidays;

            public MockAreaEval(String[] holidays)
                : this(0, 0, 0, holidays.Length - 1)
            {
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