/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is1 distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
/*
 * Created on May 30, 2005
 *
 */
namespace TestCases.HSSF.Record.Formula.Functions
{
    using System;
    using NPOI.HSSF.Record.Formula.Functions;
    using NPOI.HSSF.Record.Formula.Eval;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    [TestClass]
    public class TestStatsLib : AbstractNumericTestCase
    {
        [TestMethod]
        public void TestDevsq()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.devsq(v);
            x = 82.5;
            Assert.AreEqual(x, d, "devsq ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.devsq(v);
            x = 0;
            Assert.AreEqual(x, d, "devsq ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.devsq(v);
            x = 0;
            Assert.AreEqual(x, d, "devsq ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.devsq(v);
            x = 2.5;
            Assert.AreEqual(x, d, "devsq ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.devsq(v);
            x = 10953.7416965767;
            Assert.AreEqual(x, d,0.0000000001, "devsq ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.devsq(v);
            x = 82.5;
            Assert.AreEqual(x, d, "devsq ");
        }
        [TestMethod]
        public void TestKthLargest()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.kthLargest(v, 3);
            x = 8;
            Assert.AreEqual(x, d, "kthLargest ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.kthLargest(v, 3);
            x = 1;
            Assert.AreEqual(x, d, "kthLargest ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.kthLargest(v, 3);
            x = 0;
            Assert.AreEqual(x, d, "kthLargest ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.kthLargest(v, 3);
            x = 2;
            Assert.AreEqual(x, d, "kthLargest ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.kthLargest(v, 3);
            x = 5.37828;
            Assert.AreEqual(x, d, "kthLargest ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.kthLargest(v, 3);
            x = -3;
            Assert.AreEqual(x, d, "kthLargest ");
        }

        //public void TestKthSmallest()
        //{
        //}
        [TestMethod]
        public void TestAvedev()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.avedev(v);
            x = 2.5;
            Assert.AreEqual(x, d, "avedev ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.avedev(v);
            x = 0;
            Assert.AreEqual(x, d, "avedev ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.avedev(v);
            x = 0;
            Assert.AreEqual(x, d, "avedev ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.avedev(v);
            x = 0.5;
            Assert.AreEqual(x, d, "avedev ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.avedev(v);
            x = 36.42176053333;
            Assert.AreEqual(x, d,0.0000000001, "avedev ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.avedev(v);
            x = 2.5;
            Assert.AreEqual(x, d, "avedev ");
        }
        [TestMethod]
        public void TestMedian()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.median(v);
            x = 5.5;
            Assert.AreEqual(x, d, "median ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.median(v);
            x = 1;
            Assert.AreEqual(x, d, "median ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.median(v);
            x = 0;
            Assert.AreEqual(x, d, "median ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.median(v);
            x = 1.5;
            Assert.AreEqual(x, d, "median ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.median(v);
            x = 5.37828;
            Assert.AreEqual(x, d, "median ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.median(v);
            x = -5.5;
            Assert.AreEqual(x, d, "median ");

            v = new double[] { -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.median(v);
            x = -6;
            Assert.AreEqual(x, d, "median ");

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            d = StatsLib.median(v);
            x = 5;
            Assert.AreEqual(x, d, "median ");
        }
        [TestMethod]
        public void TestMode()
        {
            double[] v;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            ConfirmMode(v,Double.NaN);

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            ConfirmMode(v, 1.0);

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            ConfirmMode(v, 0.0);

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            ConfirmMode(v, 1.0);

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            ConfirmMode(v, Double.NaN);

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            ConfirmMode(v, Double.NaN);

            v = new double[] { 1, 2, 3, 4, 1, 1, 1, 1, 0, 0, 0, 0, 0 };
            ConfirmMode(v, 1.0);

            v = new double[] { 0, 1, 2, 3, 4, 1, 1, 1, 0, 0, 0, 0, 1 };
            ConfirmMode(v, 0.0);
        }
        private static void ConfirmMode(double[] v, double expectedResult)
        {
            double actual;
            try
            {
                actual = Mode.Evaluate(v);
                if (double.IsNaN(expectedResult))
                {
                    throw new AssertFailedException("Expected N/A exception was not thrown");
                }
            }
            catch (EvaluationException e)
            {
                if (double.IsNaN(expectedResult))
                {
                    Assert.AreEqual(ErrorEval.NA, e.GetErrorEval());
                    return;
                }
                throw;
            }
            Assert.AreEqual(expectedResult, actual, "mode");

        }
        [TestMethod]
        public void TestStddev()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.stdev(v);
            x = 3.02765035410;
            Assert.AreEqual(x, d,0.0000000001, "stdev ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.stdev(v);
            x = 0;
            Assert.AreEqual(x, d, "stdev ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.stdev(v);
            x = 0;
            Assert.AreEqual(x, d, "stdev ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.stdev(v);
            x = 0.52704627669;
            Assert.AreEqual(x, d, 0.0000000001, "stdev ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.stdev(v);
            x = 52.33006233652;
            Assert.AreEqual(x, d, 0.0000000001, "stdev ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.stdev(v);
            x = 3.02765035410;
            Assert.AreEqual(x, d, 0.0000000001, "stdev ");
        }
    }
}