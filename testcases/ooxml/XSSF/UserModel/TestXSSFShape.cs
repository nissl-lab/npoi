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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    internal class TestXSSFShape
    {
        [Test]
        public void Test58325_one()
        {
            Check58325(XSSFTestDataSamples.OpenSampleWorkbook("58325_lt.xlsx"), 1);
        }
        [Test]
        public void Test58325_three()
        {
            Check58325(XSSFTestDataSamples.OpenSampleWorkbook("58325_db.xlsx"), 3);
        }
        private void Check58325(XSSFWorkbook wb, int expectedShapes)
        {
            XSSFSheet sheet = wb.GetSheet("MetasNM001") as XSSFSheet;
            ClassicAssert.IsNotNull(sheet);
            StringBuilder str = new StringBuilder();
            str.Append("sheet " + sheet.SheetName + " - ");
            XSSFDrawing drawing = sheet.GetDrawingPatriarch();
            //drawing = ((XSSFSheet)sheet).createDrawingPatriarch();
            List<XSSFShape> shapes = drawing.GetShapes();
            str.Append("drawing.Shapes.size() = " + shapes.Count);
            IEnumerator<XSSFShape> it = shapes.GetEnumerator();
            while (it.MoveNext())
            {
                XSSFShape shape = it.Current;
                str.Append(", " + shape.ToString());
                str.Append(", Col1:" + ((XSSFClientAnchor)shape.GetAnchor()).Col1);
                str.Append(", Col2:" + ((XSSFClientAnchor)shape.GetAnchor()).Col2);
                str.Append(", Row1:" + ((XSSFClientAnchor)shape.GetAnchor()).Row1);
                str.Append(", Row2:" + ((XSSFClientAnchor)shape.GetAnchor()).Row2);
            }

            ClassicAssert.AreEqual(expectedShapes, shapes.Count,
                "Having shapes: " + str);
        }
    }
}