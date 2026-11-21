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

using NUnit.Framework;
using NUnit.Framework.Legacy;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XDDF.UserModel.Chart;

namespace TestCases.XSSF.UserModel.Charts
{
    [TestFixture]
    public class TestXDDFManualLayout
    {
        private IWorkbook wb;
        private XDDFManualLayout layout;

        [SetUp]
        public void CreateEmptyLayout()
        {
            wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            var drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            var chart = drawing.CreateChart(anchor);
            var legend = chart.GetOrAddLegend();
            layout = legend.GetOrAddManualLayout();
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

            layout.WidthRatio = (newRatio);
            ClassicAssert.IsTrue(layout.WidthRatio == newRatio);

            layout.HeightRatio = (newRatio);
            ClassicAssert.IsTrue(layout.HeightRatio == newRatio);

            layout.X= (newCoordinate);
            ClassicAssert.IsTrue(layout.X == newCoordinate);

            layout.Y= (newCoordinate);
            ClassicAssert.IsTrue(layout.Y == newCoordinate);

            layout.XMode= (nonDefaultMode);
            ClassicAssert.IsTrue(layout.XMode == nonDefaultMode);

            layout.YMode=(nonDefaultMode);
            ClassicAssert.IsTrue(layout.YMode == nonDefaultMode);

            layout.WidthMode=(nonDefaultMode);
            ClassicAssert.IsTrue(layout.WidthMode == nonDefaultMode);

            layout.HeightMode=(nonDefaultMode);
            ClassicAssert.IsTrue(layout.HeightMode == nonDefaultMode);

            layout.Target=(nonDefaultTarget);
            ClassicAssert.IsTrue(layout.Target == nonDefaultTarget);
        }

        /*
         * Layout must have reasonable default values and must not throw
         * any exceptions.
         */
        [Test]
        public void TestDefaultValues()
        {
            ClassicAssert.IsNotNull(layout.Target);
            ClassicAssert.IsNotNull(layout.XMode);
            ClassicAssert.IsNotNull(layout.YMode);
            ClassicAssert.IsNotNull(layout.HeightMode);
            ClassicAssert.IsNotNull(layout.WidthMode);
            /*
             * According to interface, 0.0 should be returned for
             * unInitialized double properties.
             */
            ClassicAssert.IsTrue(layout.X == 0.0);
            ClassicAssert.IsTrue(layout.Y == 0.0);
            ClassicAssert.IsTrue(layout.WidthRatio == 0.0);
            ClassicAssert.IsTrue(layout.HeightRatio == 0.0);
        }
    }
}
