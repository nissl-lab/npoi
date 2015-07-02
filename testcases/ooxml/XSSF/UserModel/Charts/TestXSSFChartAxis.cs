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

using System;
using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
namespace NPOI.XSSF.UserModel.Charts
{
    [TestFixture]
    public class TestXSSFChartAxis
    {

        private static double EPSILON = 1E-7;
        private IChartAxis axis;

        public TestXSSFChartAxis()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IDrawing Drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            IChart chart = Drawing.CreateChart(anchor);
            axis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Bottom);
        }
        [Test]
        public void TestLogBaseIllegalArgument()
        {
            ArgumentException iae = null;
            try
            {
                axis.LogBase=(0.0);
            }
            catch (ArgumentException e)
            {
                iae = e;
            }
            Assert.IsNotNull(iae);

            iae = null;
            try
            {
                axis.LogBase=(30000.0);
            }
            catch (ArgumentException e)
            {
                iae = e;
            }
            Assert.IsNotNull(iae);
        }
        [Test]
        public void TestLogBaseLegalArgument()
        {
            axis.LogBase=(Math.E);
            Assert.IsTrue(Math.Abs(axis.LogBase - Math.E) < EPSILON);
        }
        [Test]
        public void TestNumberFormat()
        {
            String numberFormat = "General";
            axis.NumberFormat=(numberFormat);
            Assert.AreEqual(numberFormat, axis.NumberFormat);
        }
        [Test]
        public void TestMaxAndMinAccessMethods()
        {
            double newValue = 10.0;

            axis.Minimum=(newValue);
            Assert.IsTrue(Math.Abs(axis.Minimum - newValue) < EPSILON);

            axis.Maximum=(newValue);
            Assert.IsTrue(Math.Abs(axis.Maximum - newValue) < EPSILON);
        }
        [Test]
        public void TestVisibleAccessMethods()
        {
            axis.IsVisible = (/*setter*/true);
            Assert.IsTrue(axis.IsVisible);

            axis.IsVisible = (/*setter*/false);
            Assert.IsFalse(axis.IsVisible);
        }

        [Test]
        public void TestMajorTickMarkAccessMethods()
        {
            axis.MajorTickMark = (/*setter*/AxisTickMark.None);
            Assert.AreEqual(AxisTickMark.None, axis.MajorTickMark);

            axis.MajorTickMark = (/*setter*/AxisTickMark.In);
            Assert.AreEqual(AxisTickMark.In, axis.MajorTickMark);

            axis.MajorTickMark = (/*setter*/AxisTickMark.Out);
            Assert.AreEqual(AxisTickMark.Out, axis.MajorTickMark);

            axis.MajorTickMark = (/*setter*/AxisTickMark.Cross);
            Assert.AreEqual(AxisTickMark.Cross, axis.MajorTickMark);
        }

        [Test]
        public void TestMinorTickMarkAccessMethods()
        {
            axis.MinorTickMark = (/*setter*/AxisTickMark.None);
            Assert.AreEqual(AxisTickMark.None, axis.MinorTickMark);

            axis.MinorTickMark = (/*setter*/AxisTickMark.In);
            Assert.AreEqual(AxisTickMark.In, axis.MinorTickMark);

            axis.MinorTickMark = (/*setter*/AxisTickMark.Out);
            Assert.AreEqual(AxisTickMark.Out, axis.MinorTickMark);

            axis.MinorTickMark = (/*setter*/AxisTickMark.Cross);
            Assert.AreEqual(AxisTickMark.Cross, axis.MinorTickMark);
        }

    }
}