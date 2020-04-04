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
    using NUnit.Framework;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NPOI.SS.Formula.Functions;

    /**
     * @author Josh Micich
     */
    [TestFixture]
    public class TestDays360
    {

        /**
         * @param month 1-based
         */
        private static DateTime MakeDate(int year, int month, int day)
        {

            //Calendar cal = new GregorianCalendar(year, month-1, day, 0, 0, 0);
            //cal.Set(Calendar.MILLISECOND, 0);
            //return cal.GetTime();
            return new DateTime(year, month, day, 0, 0, 0, 0);
        }
        private static DateTime decrementDay(DateTime d)
        {
            //Calendar c = new GregorianCalendar();
            //c.SetTimeInMillis(d.GetTime());
            //c.Add(Calendar.DAY_OF_MONTH, -1);
            //return c.GetTime();
            return d.AddDays(-1);
        }
        private static String fmt(DateTime d)
        {
            //Calendar c = new GregorianCalendar();
            //c.SetTimeInMillis(d.GetTime());
            //StringBuilder sb = new StringBuilder();
            //sb.Append(c.Get(Calendar.YEAR));
            //sb.Append("/");
            //sb.Append(c.Get(Calendar.MONTH)+1);
            //sb.Append("/");
            //sb.Append(c.Get(Calendar.DAY_OF_MONTH));
            //return sb.ToString();
            return d.ToString("yyyy/mm/dd");
        }

        [Test]
        public void TestBasic()
        {
            Confirm(120, 2009, 1, 15, 2009, 5, 15);
            Confirm(158, 2009, 1, 26, 2009, 7, 4);

            // same results in leap years
            Confirm(120, 2008, 1, 15, 2008, 5, 15);
            Confirm(158, 2008, 1, 26, 2008, 7, 4);

            // longer time spans
            Confirm(562, 2008, 8, 11, 2010, 3, 3);
            Confirm(916, 2007, 2, 23, 2009, 9, 9);
        }

        private static void Confirm(int expResult, int y1, int m1, int d1, int y2, int m2, int d2)
        {
            Confirm(expResult, MakeDate(y1, m1, d1), MakeDate(y2, m2, d2), false);
            Confirm(-expResult, MakeDate(y2, m2, d2), MakeDate(y1, m1, d1), false);

            // other tests
            Confirm(1, MakeDate(1993, 2, 28), MakeDate(1993, 3, 1), false);
            Confirm(1, MakeDate(1996, 2, 29), MakeDate(1996, 3, 1), false);
            Confirm(-2, MakeDate(1993, 2, 28), MakeDate(1993, 2, 28), false);
            Confirm(3, MakeDate(1993, 2, 28), MakeDate(1993, 3, 1), true);
            Confirm(2, MakeDate(1996, 2, 29), MakeDate(1996, 3, 1), true);
        }
        /**
         * The <c>method</c> parameter only Makes a difference when the second parameter
         * is the last day of the month that does <em>not</em> have 30 days.
         */
        [Test]
        public void TestMonthBoundaries()
        {
            // jan
            ConfirmMonthBoundary(false, 2001, 1, 0, 0, 2, 3, 4);
            ConfirmMonthBoundary(true, 2001, 1, 0, 0, 1, 2, 3);
            // feb
            ConfirmMonthBoundary(false, 2001, 2, -2, 1, 2, 3, 4);
            ConfirmMonthBoundary(true, 2001, 2, 0, 1, 2, 3, 4);
            // mar
            ConfirmMonthBoundary(false, 2001, 3, 0, 0, 2, 3, 4);
            ConfirmMonthBoundary(true, 2001, 3, 0, 0, 1, 2, 3);
            // apr
            ConfirmMonthBoundary(false, 2001, 4, 0, 1, 2, 3, 4);
            ConfirmMonthBoundary(true, 2001, 4, 0, 1, 2, 3, 4);
            // may
            ConfirmMonthBoundary(false, 2001, 5, 0, 0, 2, 3, 4);
            ConfirmMonthBoundary(true, 2001, 5, 0, 0, 1, 2, 3);
            // jun
            ConfirmMonthBoundary(false, 2001, 6, 0, 1, 2, 3, 4);
            ConfirmMonthBoundary(true, 2001, 6, 0, 1, 2, 3, 4);
            // leap year
            ConfirmMonthBoundary(false, 2012, 2, -1, 1, 2, 3, 4);
            ConfirmMonthBoundary(true, 2012, 2, 0, 1, 2, 3, 4);
        }


        /**
         * @param monthNo 1-based
         * @param diffs
         */
        private static void ConfirmMonthBoundary(bool method, int year, int monthNo, params int[] diffs)
        {
            DateTime firstDayOfNextMonth = MakeDate(year, monthNo + 1, 1);
            DateTime secondArg = decrementDay(firstDayOfNextMonth);
            DateTime firstArg = secondArg;

            foreach (int expResult in diffs)
            {
                Confirm(expResult, firstArg, secondArg, method);
                firstArg = decrementDay(firstArg);
            }

        }
        private static void Confirm(int expResult, DateTime firstArg, DateTime secondArg, bool method)
        {

            ValueEval ve;
            if (method)
            {
                // TODO enable 3rd arg -
                ve = invokeDays360(Convert(firstArg), Convert(secondArg), BoolEval.ValueOf(method));
            }
            else
            {
                ve = invokeDays360(Convert(firstArg), Convert(secondArg));
            }
            if (ve is NumberEval)
            {
                Assert.IsTrue(ve is NumberEval, "wrong return type (" + ve.GetType().Name + ")");
                NumberEval numberEval = (NumberEval)ve;
                String err = String.Format("days360({0},{1},{2}) wrong result", firstArg, secondArg, method);
                Assert.AreEqual(expResult, numberEval.NumberValue, err);
            }
        }
        private static ValueEval invokeDays360(params ValueEval[] args)
        {
            return new Days360().Evaluate(args, -1, -1);
        }
        private static NumberEval Convert(DateTime d)
        {
            return new NumberEval(DateUtil.GetExcelDate(d));
        }
    }


}