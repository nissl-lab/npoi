/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
/*
 * Created on May 30, 2005
 *
 */
namespace TestCases.SS.Formula.Functions
{

    using NPOI.SS.Formula.Eval;
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;
    using System;


    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *
     */
    [TestFixture]
    public class TestStatsLib : AbstractNumericTestCase
    {
        [Test]
        public void TestDevsq()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.devsq(v);
            x = 82.5;
            Assert.AreEqual( x, d,"devsq ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.devsq(v);
            x = 0;
            Assert.AreEqual( x, d,"devsq ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.devsq(v);
            x = 0;
            Assert.AreEqual( x, d,"devsq ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.devsq(v);
            x = 2.5;
            Assert.AreEqual( x, d,"devsq ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.devsq(v);
            x = 10953.7416965767;
            Assert.AreEqual( x, d, 0.0000000001, "devsq ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.devsq(v);
            x = 82.5;
            Assert.AreEqual( x, d,"devsq ");
        }
        [Test]
        public void TestKthLargest()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.kthLargest(v, 3);
            x = 8;
            Assert.AreEqual( x, d,"kthLargest ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.kthLargest(v, 3);
            x = 1;
            Assert.AreEqual( x, d,"kthLargest ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.kthLargest(v, 3);
            x = 0;
            Assert.AreEqual( x, d,"kthLargest ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.kthLargest(v, 3);
            x = 2;
            Assert.AreEqual( x, d,"kthLargest ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.kthLargest(v, 3);
            x = 5.37828;
            Assert.AreEqual( x, d,"kthLargest ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.kthLargest(v, 3);
            x = -3;
            Assert.AreEqual( x, d,"kthLargest ");
        }
        [Test]
        public void TestKthSmallest()
        {
        }
        [Test]
        public void TestAvedev()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.avedev(v);
            x = 2.5;
            Assert.AreEqual( x, d,"avedev ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.avedev(v);
            x = 0;
            Assert.AreEqual( x, d,"avedev ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.avedev(v);
            x = 0;
            Assert.AreEqual( x, d,"avedev ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.avedev(v);
            x = 0.5;
            Assert.AreEqual( x, d,"avedev ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.avedev(v);
            x = 36.42176053333;
            Assert.AreEqual( x, d,0.00000000001,"avedev ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.avedev(v);
            x = 2.5;
            Assert.AreEqual( x, d,"avedev ");
        }
        [Test]
        public void TestMedian()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.median(v);
            x = 5.5;
            Assert.AreEqual( x, d,"median ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.median(v);
            x = 1;
            Assert.AreEqual( x, d,"median ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.median(v);
            x = 0;
            Assert.AreEqual( x, d,"median ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.median(v);
            x = 1.5;
            Assert.AreEqual( x, d,"median ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.median(v);
            x = 5.37828;
            Assert.AreEqual( x, d,"median ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.median(v);
            x = -5.5;
            Assert.AreEqual( x, d,"median ");

            v = new double[] { -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.median(v);
            x = -6;
            Assert.AreEqual( x, d,"median ");

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            d = StatsLib.median(v);
            x = 5;
            Assert.AreEqual( x, d,"median ");
        }
        [Test]
        public void TestMode()
        {
            double[] v;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            ConfirmMode(v, null);

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            ConfirmMode(v, 1.0);

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            ConfirmMode(v, 0.0);

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            ConfirmMode(v, 1.0);

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            ConfirmMode(v, null);

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            ConfirmMode(v, null);

            v = new double[] { 1, 2, 3, 4, 1, 1, 1, 1, 0, 0, 0, 0, 0 };
            ConfirmMode(v, 1.0);

            v = new double[] { 0, 1, 2, 3, 4, 1, 1, 1, 0, 0, 0, 0, 1 };
            ConfirmMode(v, 0.0);
        }
        private static void ConfirmMode(double[] v, double expectedResult)
        {
            ConfirmMode(v, (Double?)expectedResult);
        }
        private static void ConfirmMode(double[] v, Double? expectedResult)
        {
            double actual;
            try
            {
                actual = Mode.Evaluate(v);
                if (expectedResult == null)
                {
                    throw new AssertionException("Expected N/A exception was not thrown");
                }
            }
            catch (EvaluationException e)
            {
                if (expectedResult == null)
                {
                    Assert.AreEqual(ErrorEval.NA, e.GetErrorEval());
                    return;
                }
                throw e;
            }
            Assert.AreEqual( expectedResult.Value, actual,"mode");
        }

        [Test]
        public void TestStddev()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            d = StatsLib.stdev(v);
            x = 3.02765035409749;
            Assert.AreEqual( x, d,0.0000000001, "stdev ");

            v = new double[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
            d = StatsLib.stdev(v);
            x = 0;
            Assert.AreEqual( x, d,"stdev ");

            v = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            d = StatsLib.stdev(v);
            x = 0;
            Assert.AreEqual( x, d,"stdev ");

            v = new double[] { 1, 2, 1, 2, 1, 2, 1, 2, 1, 2 };
            d = StatsLib.stdev(v);
            x = 0.52704627669;
            Assert.AreEqual( x, d,0.000000001,"stdev ");

            v = new double[] { 123.12, 33.3333, 2d / 3d, 5.37828, 0.999 };
            d = StatsLib.stdev(v);
            x = 52.33006233652;
            Assert.AreEqual( x, d,0.0000000001,"stdev ");

            v = new double[] { -1, -2, -3, -4, -5, -6, -7, -8, -9, -10 };
            d = StatsLib.stdev(v);
            x = 3.02765035410;
            Assert.AreEqual( x, d,0.0000000001,"stdev ");
        }
        [Test]
        
        public void TestVar()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 3.50, 5.00, 7.23, 2.99 };
            d = StatsLib.var(v);
            x = 3.6178;
            //the following AreEqual add a delta param against java version, otherwise tests fail.
            Assert.AreEqual( x, d, 0.00001,"var ");

            v = new double[] { 34.5, 2.0, 8.9, -4.0 };
            d = StatsLib.var(v);
            x = 286.99;
            Assert.AreEqual( x, d,0.001,"var ");

            v = new double[] { 7.0, 25.0, 21.69 };
            d = StatsLib.var(v);
            x = 91.79203333;
            Assert.AreEqual( x, d,0.00000001,"var ");

            v = new double[] { 1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299 };
            d = StatsLib.var(v);
            x = 754.2666667;
            Assert.AreEqual( x, d, 0.0000001,"var ");
        }
        [Test]
        public void TestVarp()
        {
            double[] v = null;
            double d, x = 0;

            v = new double[] { 3.50, 5.00, 7.23, 2.99 };
            d = StatsLib.varp(v);
            x = 2.71335;
            Assert.AreEqual( x, d, 0.000001, "varp ");

            v = new double[] { 34.5, 2.0, 8.9, -4.0 };
            d = StatsLib.varp(v);
            x = 215.2425;
            Assert.AreEqual( x, d,0.00001,"varp ");

            v = new double[] { 7.0, 25.0, 21.69 };
            d = StatsLib.varp(v);
            x = 61.19468889;
            Assert.AreEqual( x, d, 0.00000001, "varp ");

            v = new double[] { 1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299 };
            d = StatsLib.varp(v);
            x = 678.84;
            Assert.AreEqual( x, d, 0.001,"varp ");
        }
    }

}