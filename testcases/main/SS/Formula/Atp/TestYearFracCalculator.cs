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

namespace TestCases.SS.Formula.Atp
{


    using System;
    using NUnit.Framework;
    using NPOI.SS.Formula.Atp;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;

    /**
     * Specific Test cases for YearFracCalculator
     */
    [TestFixture]
    public class TestYearFracCalculator
    {
        [Test]
        public void TestBasis1()
        {
            Confirm(md(1999, 1, 1), md(1999, 4, 5), 1, 0.257534247);
            Confirm(md(1999, 4, 1), md(1999, 4, 5), 1, 0.010958904);
            Confirm(md(1999, 4, 1), md(1999, 4, 4), 1, 0.008219178);
            Confirm(md(1999, 4, 2), md(1999, 4, 5), 1, 0.008219178);
            Confirm(md(1999, 3, 31), md(1999, 4, 3), 1, 0.008219178);
            Confirm(md(1999, 4, 5), md(1999, 4, 8), 1, 0.008219178);
            Confirm(md(1999, 4, 4), md(1999, 4, 7), 1, 0.008219178);
            Confirm(md(2000, 2, 5), md(2000, 6, 1), 0, 0.322222222);
        }

        private void Confirm(double startDate, double endDate, int basis, double expectedValue)
        {
            double actualValue;
            try
            {
                actualValue = YearFracCalculator.Calculate(startDate, endDate, basis);
            }
            catch (EvaluationException e)
            {
                throw e;
            }
            double diff = actualValue - expectedValue;
            if (Math.Abs(diff) > 0.000000001)
            {
                double hours = diff * 365 * 24;
                Console.WriteLine(startDate + " " + endDate + " off by " + hours + " hours");
                Assert.AreEqual(expectedValue, actualValue, 0.000000001);
            }

        }

        private static double md(int year, int month, int day)
        {
            DateTime d = new DateTime(year, month, day, 0, 0, 0, 0);
            return DateUtil.GetExcelDate(d);
        }
    }

}