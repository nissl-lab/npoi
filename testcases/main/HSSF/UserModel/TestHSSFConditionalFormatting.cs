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

namespace TestCases.HSSF.UserModel
{
    using System;

    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Util;
    using NPOI.HSSF.Record;
    using NUnit.Framework;
    using NPOI.SS.Util;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.Record.CF;
    using TestCases.SS.UserModel;

    /**
* 
* @author Dmitriy Kumshayev
*/
    [TestFixture]
    public class TestHSSFConditionalFormatting : BaseTestConditionalFormatting
    {
        protected override void AssertColour(String hexExpected, IColor actual)
        {
            Assert.IsNotNull(actual, "Colour must be given");

            if (actual is HSSFColor) {
                HSSFColor colour = (HSSFColor)actual;
                Assert.AreEqual(hexExpected, colour.GetHexString());
            } else {
                HSSFExtendedColor colour = (HSSFExtendedColor)actual;
                if (hexExpected.Length == 8)
                {
                    Assert.AreEqual(hexExpected, colour.ARGBHex);
                }
                else
                {
                    Assert.AreEqual(hexExpected, colour.ARGBHex.Substring(2));
                }
            }
        }
        [Test]
        public void TestReadOffice2007()
        {
            TestReadOffice2007("NewStyleConditionalFormattings.xls");
        }

        [Test]
        public void Test53691()
        {
            ISheetConditionalFormatting cf;
            IWorkbook wb = HSSFITestDataProvider.Instance.OpenSampleWorkbook("53691.xls");
            /*
            FileInputStream s = new FileInputStream("C:\\temp\\53691bbadfixed.xls");
            try {
                wb = new HSSFWorkbook(s);
            } finally {
                s.Close();
            }

            wb.RemoveSheetAt(1);*/

            // Initially it is good
            WriteTemp53691(wb, "agood");

            // clone sheet corrupts it
            ISheet sheet = wb.CloneSheet(0);
            WriteTemp53691(wb, "bbad");

            // removing the sheet Makes it good again
            wb.RemoveSheetAt(wb.GetSheetIndex(sheet));
            WriteTemp53691(wb, "cgood");

            // cloning again and removing the conditional formatting Makes it good again
            sheet = wb.CloneSheet(0);
            RemoveConditionalFormatting(sheet);
            WriteTemp53691(wb, "dgood");

            // cloning the conditional formatting manually Makes it bad again
            cf = sheet.SheetConditionalFormatting;
            ISheetConditionalFormatting scf = wb.GetSheetAt(0).SheetConditionalFormatting;
            for (int j = 0; j < scf.NumConditionalFormattings; j++)
            {
                cf.AddConditionalFormatting(scf.GetConditionalFormattingAt(j));
            }
            WriteTemp53691(wb, "ebad");

            // remove all conditional formatting for comparing BIFF output
            RemoveConditionalFormatting(sheet);
            RemoveConditionalFormatting(wb.GetSheetAt(0));
            WriteTemp53691(wb, "fgood");

            wb.Close();
        }

        private void RemoveConditionalFormatting(ISheet sheet)
        {
            ISheetConditionalFormatting cf = sheet.SheetConditionalFormatting;
            for (int j = 0; j < cf.NumConditionalFormattings; j++)
            {
                cf.RemoveConditionalFormatting(j);
            }
        }

        private void WriteTemp53691(IWorkbook wb, String suffix)
        {
            // assert that we can Write/read it in memory
            IWorkbook wbBack = HSSFITestDataProvider.Instance.WriteOutAndReadBack(wb);
            Assert.IsNotNull(wbBack);
            wbBack.Close();
        }
    }
}