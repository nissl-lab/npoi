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

using NUnit.Framework;using NUnit.Framework.Legacy;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.XSSF.UserModel;

namespace TestCases.XSSF.UserModel.Charts
{
    [TestFixture]
    public class TestXSSFManualLayout
    {
        private IWorkbook wb;
        private IManualLayout layout;

        [SetUp]
        public void CreateEmptyLayout()
        {
            wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IDrawing<IShape> drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            IChart chart = drawing.CreateChart(anchor);
            IChartLegend legend = chart.GetOrCreateLegend();
            layout = legend.GetManualLayout();
        }

        [TearDown]
        public void CloseWB()
        {
            wb.Close();
        }

        /*
         * Accessor methods are not trivial. They use lazy underlying bean
         * Initialization so there can be some errors (NPE, for example).
         */
        [Test]
        public void TestAccessorMethods()
        {
            double newRatio = 1.1;
            double newCoordinate = 0.3;
            LayoutMode nonDefaultMode = LayoutMode.Factor;
            LayoutTarget nonDefaultTarget = LayoutTarget.Outer;

            layout.SetWidthRatio(newRatio);
            ClassicAssert.IsTrue(layout.GetWidthRatio() == newRatio);

            layout.SetHeightRatio(newRatio);
            ClassicAssert.IsTrue(layout.GetHeightRatio() == newRatio);

            layout.SetX(newCoordinate);
            ClassicAssert.IsTrue(layout.GetX() == newCoordinate);

            layout.SetY(newCoordinate);
            ClassicAssert.IsTrue(layout.GetY() == newCoordinate);

            layout.SetXMode(nonDefaultMode);
            ClassicAssert.IsTrue(layout.GetXMode() == nonDefaultMode);

            layout.SetYMode(nonDefaultMode);
            ClassicAssert.IsTrue(layout.GetYMode() == nonDefaultMode);

            layout.SetWidthMode(nonDefaultMode);
            ClassicAssert.IsTrue(layout.GetWidthMode() == nonDefaultMode);

            layout.SetHeightMode(nonDefaultMode);
            ClassicAssert.IsTrue(layout.GetHeightMode() == nonDefaultMode);

            layout.SetTarget(nonDefaultTarget);
            ClassicAssert.IsTrue(layout.GetTarget() == nonDefaultTarget);
        }

        /*
         * Layout must have reasonable default values and must not throw
         * any exceptions.
         */
        [Test]
        public void TestDefaultValues()
        {
            ClassicAssert.IsNotNull(layout.GetTarget());
            ClassicAssert.IsNotNull(layout.GetXMode());
            ClassicAssert.IsNotNull(layout.GetYMode());
            ClassicAssert.IsNotNull(layout.GetHeightMode());
            ClassicAssert.IsNotNull(layout.GetWidthMode());
            /*
             * According to interface, 0.0 should be returned for
             * unInitialized double properties.
             */
            ClassicAssert.IsTrue(layout.GetX() == 0.0);
            ClassicAssert.IsTrue(layout.GetY() == 0.0);
            ClassicAssert.IsTrue(layout.GetWidthRatio() == 0.0);
            ClassicAssert.IsTrue(layout.GetHeightRatio() == 0.0);
        }
    }
}
