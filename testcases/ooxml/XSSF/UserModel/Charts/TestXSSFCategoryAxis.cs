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

namespace NPOI.XSSF.UserModel.Charts
{
    using System;

    using NUnit.Framework;

    using NPOI.SS.UserModel.Charts;
    using NPOI.XSSF.UserModel;

    [TestFixture]
    public class TestXSSFCategoryAxis
    {

        [Test]
        public void TestAccessMethods()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.CreateSheet() as XSSFSheet;
            XSSFDrawing Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            XSSFClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30) as XSSFClientAnchor;
            XSSFChart chart = Drawing.CreateChart(anchor) as XSSFChart;
            XSSFCategoryAxis axis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom) as XSSFCategoryAxis;

            axis.Crosses=(AxisCrosses.AutoZero);
            Assert.AreEqual(axis.Crosses, AxisCrosses.AutoZero);

            Assert.AreEqual(chart.GetAxis().Count, 1);
        }
    }

}