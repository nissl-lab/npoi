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
using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;
using NPOI.XSSF.UserModel.Extensions;

namespace TestCases.XSSF.UserModel.Extensions
{

    [TestFixture]
    public class TestXSSFBorder
    {

        [Test]
        public void TestGetBorderStyle()
        {
            CT_Stylesheet stylesheet = new CT_Stylesheet();
            CT_Border border = stylesheet.AddNewBorders().AddNewBorder();
            CT_BorderPr top = border.AddNewTop();
            CT_BorderPr right = border.AddNewRight();
            CT_BorderPr bottom = border.AddNewBottom();

            top.style = (ST_BorderStyle.dashDot);
            right.style = (ST_BorderStyle.none);
            bottom.style = (ST_BorderStyle.thin);

            XSSFCellBorder cellBorderStyle = new XSSFCellBorder(border);
            Assert.AreEqual("DashDot", cellBorderStyle.GetBorderStyle(BorderSide.TOP).ToString());

            Assert.AreEqual("None", cellBorderStyle.GetBorderStyle(BorderSide.RIGHT).ToString());
            Assert.AreEqual(BorderStyle.None, cellBorderStyle.GetBorderStyle(BorderSide.RIGHT));

            Assert.AreEqual("Thin", cellBorderStyle.GetBorderStyle(BorderSide.BOTTOM).ToString());

            Assert.AreEqual(BorderStyle.Thin, cellBorderStyle.GetBorderStyle(BorderSide.BOTTOM));
        }

    }
}
