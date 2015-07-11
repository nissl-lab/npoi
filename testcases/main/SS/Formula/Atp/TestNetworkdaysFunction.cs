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
using System.Globalization;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Atp;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Util;
using NUnit.Framework;

namespace TestCases.SS.Formula.Atp
{
    [TestFixture]
    public class TestNetworkdaysFunction
    {
        private static int OCTOBER = 10;
        private static int NOVEMBER = 11;
        private static int DECEMBER = 12;
        private static int JANUARY = 1;
        private static int MARCH = 3;
        private static SimpleDateFormat formatter = new SimpleDateFormat("yyyy/MM/dd");

        private static String STARTING_DATE = formatter.Format(new DateTime(2008, OCTOBER, 1), CultureInfo.CurrentCulture);

        private static String END_DATE = formatter.Format(new DateTime(2009, MARCH, 1), CultureInfo.CurrentCulture);

        private static String FIRST_HOLIDAY = formatter.Format(new DateTime(2008, NOVEMBER, 26), CultureInfo.CurrentCulture);

        private static String SECOND_HOLIDAY = formatter.Format(new DateTime(2008, DECEMBER, 4), CultureInfo.CurrentCulture);

        private static String THIRD_HOLIDAY = formatter.Format(new DateTime(2009, JANUARY, 21), CultureInfo.CurrentCulture);

        private static OperationEvaluationContext EC = new OperationEvaluationContext(null, null, 1, 1, 1, null);
        [Test]
        public void TestFailWhenNoArguments()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, NetworkdaysFunction.instance.Evaluate(new ValueEval[0], null));
        }
        [Test]
        public void TestFailWhenLessThan2Arguments()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, NetworkdaysFunction.instance.Evaluate(new ValueEval[1], null));
        }
        [Test]
        public void TestFailWhenMoreThan3Arguments()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, NetworkdaysFunction.instance.Evaluate(new ValueEval[4], null));
        }
        [Test]
        public void TestFailWhenArgumentsAreNotDates()
        {
            Assert.AreEqual(ErrorEval.VALUE_INVALID, NetworkdaysFunction.instance.Evaluate(new ValueEval[]{ new StringEval("Potato"),
                new StringEval("Cucumber") }, EC));
        }
        [Test]
        public void TestFailWhenStartDateAfterEndDate()
        {
            Assert.AreEqual(ErrorEval.NAME_INVALID, NetworkdaysFunction.instance.Evaluate(new ValueEval[]{ new StringEval(END_DATE),
                new StringEval(STARTING_DATE) }, EC));
        }
        [Test]
        public void TestReturnNetworkdays()
        {
            Assert.AreEqual(108, (int)((NumericValueEval)NetworkdaysFunction.instance.Evaluate(new ValueEval[]{
                new StringEval(STARTING_DATE), new StringEval(END_DATE) }, EC)).NumberValue);
        }
        [Test]
        public void TestReturnNetworkdaysWithAHoliday()
        {
            Assert.AreEqual(107, (int)((NumericValueEval)NetworkdaysFunction.instance.Evaluate(new ValueEval[]{
                new StringEval(STARTING_DATE), new StringEval(END_DATE), new StringEval(FIRST_HOLIDAY) },
                    EC)).NumberValue);
        }
        [Test]
        public void TestReturnNetworkdaysWithManyHolidays()
        {
            Assert.AreEqual(105, (int)((NumericValueEval)NetworkdaysFunction.instance.Evaluate(new ValueEval[]{
                new StringEval(STARTING_DATE), new StringEval(END_DATE),
                new MockAreaEval(new string[]{FIRST_HOLIDAY, SECOND_HOLIDAY, THIRD_HOLIDAY}) }, EC)).NumberValue);
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