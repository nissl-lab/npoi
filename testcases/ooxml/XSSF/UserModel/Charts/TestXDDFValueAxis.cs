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

using NPOI.SS.UserModel;
using NPOI.XDDF.UserModel.Chart;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace TestCases.XSSF.UserModel.Charts
{
    [TestFixture]
    public class TestXDDFValueAxis
    {
        [Test]
        public void TestAccessMethods()
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            var Drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 1, 1, 10, 30);
            var chart = Drawing.CreateChart(anchor);
            var axis = chart.CreateValueAxis(AxisPosition.Bottom);

            axis.CrossBetween = (AxisCrossBetween.MidpointCategory);
            ClassicAssert.AreEqual(axis.CrossBetween, AxisCrossBetween.MidpointCategory);

            axis.Crosses=(AxisCrosses.AutoZero);
            ClassicAssert.AreEqual(axis.Crosses, AxisCrosses.AutoZero);

            ClassicAssert.AreEqual(chart.GetAxis().Count, 1);
        }
    }
}

