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

   2012 - Alfresco Software, Ltd.
   Alfresco Software has modified source of this file
   The details of Changes as svn diff can be found in svn at location root/projects/3rd-party/src 
==================================================================== */

namespace TestCases.SS.UserModel
{
    using System;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using TestCases.SS;

    /**
     * Tests of {@link BorderStyle}
     */
    public abstract class BaseTestBorderStyle
    {

        private ITestDataProvider _testDataProvider;

        protected BaseTestBorderStyle(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }

        /**
         * Test that we use the specified locale when deciding
         *   how to format normal numbers
         */
        [Test]
        public void TestBorderStyle()
        {
            String ext = _testDataProvider.StandardFileNameExtension;
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook("59264." + ext);
            ISheet sh = wb.GetSheetAt(0);

            AssertBorderStyleEquals(BorderStyle.None, GetDiagonalCell(sh, 0));
            AssertBorderStyleEquals(BorderStyle.Thin, GetDiagonalCell(sh, 1));
            AssertBorderStyleEquals(BorderStyle.Medium, GetDiagonalCell(sh, 2));
            AssertBorderStyleEquals(BorderStyle.Dashed, GetDiagonalCell(sh, 3));
            AssertBorderStyleEquals(BorderStyle.Dotted, GetDiagonalCell(sh, 4));
            AssertBorderStyleEquals(BorderStyle.Thick, GetDiagonalCell(sh, 5));
            AssertBorderStyleEquals(BorderStyle.Double, GetDiagonalCell(sh, 6));
            AssertBorderStyleEquals(BorderStyle.Hair, GetDiagonalCell(sh, 7));
            AssertBorderStyleEquals(BorderStyle.MediumDashed, GetDiagonalCell(sh, 8));
            AssertBorderStyleEquals(BorderStyle.DashDot, GetDiagonalCell(sh, 9));
            AssertBorderStyleEquals(BorderStyle.MediumDashDot, GetDiagonalCell(sh, 10));
            AssertBorderStyleEquals(BorderStyle.DashDotDot, GetDiagonalCell(sh, 11));
            AssertBorderStyleEquals(BorderStyle.MediumDashDotDot, GetDiagonalCell(sh, 12));
            AssertBorderStyleEquals(BorderStyle.SlantedDashDot, GetDiagonalCell(sh, 13));

            wb.Close();
        }

        private ICell GetDiagonalCell(ISheet sheet, int n)
        {
            return sheet.GetRow(n).GetCell(n);
        }

        protected void AssertBorderStyleEquals(BorderStyle expected, ICell cell)
        {
            ICellStyle style = cell.CellStyle;
            Assert.AreEqual(expected, style.BorderTop);
            Assert.AreEqual(expected, style.BorderBottom);
            Assert.AreEqual(expected, style.BorderLeft);
            Assert.AreEqual(expected, style.BorderRight);
        }

    }
}
