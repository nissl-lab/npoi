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
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCases.SS.Util
{
    [TestFixture]
    public class TestRegionUtil
    {
        private static CellRangeAddress A1C3 = new CellRangeAddress(0, 2, 0, 2);
        private static BorderStyle NONE = BorderStyle.None;
        private static BorderStyle THIN = BorderStyle.Thin;
        private static int RED = IndexedColors.Red.Index;
        private static int DEFAULT_COLOR = 0;
        private IWorkbook wb;
        private ISheet sheet;

        [SetUp]
        public void SetUp()
        {
            wb = new HSSFWorkbook();
            sheet = wb.CreateSheet();
        }

        [TearDown]
        public void TearDown()
        {
            wb.Close();
        }

        private ICellStyle GetCellStyle(int rowIndex, int columnIndex)
        {
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
            return cell.CellStyle;
        }

        [Test]
        public void SetBorderTop()
        {
            Assert.AreEqual(NONE, GetCellStyle(0, 0).BorderTop);
            Assert.AreEqual(NONE, GetCellStyle(0, 1).BorderTop);
            Assert.AreEqual(NONE, GetCellStyle(0, 2).BorderTop);
            RegionUtil.SetBorderTop(THIN, A1C3, sheet);
            Assert.AreEqual(THIN, GetCellStyle(0, 0).BorderTop);
            Assert.AreEqual(THIN, GetCellStyle(0, 1).BorderTop);
            Assert.AreEqual(THIN, GetCellStyle(0, 2).BorderTop);
        }
        [Test]
        public void SetBorderBottom()
        {
            Assert.AreEqual(NONE, GetCellStyle(2, 0).BorderBottom);
            Assert.AreEqual(NONE, GetCellStyle(2, 1).BorderBottom);
            Assert.AreEqual(NONE, GetCellStyle(2, 2).BorderBottom);
            RegionUtil.SetBorderBottom(THIN, A1C3, sheet);
            Assert.AreEqual(THIN, GetCellStyle(2, 0).BorderBottom);
            Assert.AreEqual(THIN, GetCellStyle(2, 1).BorderBottom);
            Assert.AreEqual(THIN, GetCellStyle(2, 2).BorderBottom);
        }
        [Test]
        public void SetBorderRight()
        {
            Assert.AreEqual(NONE, GetCellStyle(0, 2).BorderRight);
            Assert.AreEqual(NONE, GetCellStyle(1, 2).BorderRight);
            Assert.AreEqual(NONE, GetCellStyle(2, 2).BorderRight);
            RegionUtil.SetBorderRight(THIN, A1C3, sheet);
            Assert.AreEqual(THIN, GetCellStyle(0, 2).BorderRight);
            Assert.AreEqual(THIN, GetCellStyle(1, 2).BorderRight);
            Assert.AreEqual(THIN, GetCellStyle(2, 2).BorderRight);
        }
        [Test]
        public void SetBorderLeft()
        {
            Assert.AreEqual(NONE, GetCellStyle(0, 0).BorderLeft);
            Assert.AreEqual(NONE, GetCellStyle(1, 0).BorderLeft);
            Assert.AreEqual(NONE, GetCellStyle(2, 0).BorderLeft);
            RegionUtil.SetBorderLeft(THIN, A1C3, sheet);
            Assert.AreEqual(THIN, GetCellStyle(0, 0).BorderLeft);
            Assert.AreEqual(THIN, GetCellStyle(1, 0).BorderLeft);
            Assert.AreEqual(THIN, GetCellStyle(2, 0).BorderLeft);
        }

        [Test]
        public void SetTopBorderColor()
        {
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(0, 0).TopBorderColor);
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(0, 1).TopBorderColor);
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(0, 2).TopBorderColor);
            RegionUtil.SetTopBorderColor(RED, A1C3, sheet);
            Assert.AreEqual(RED, GetCellStyle(0, 0).TopBorderColor);
            Assert.AreEqual(RED, GetCellStyle(0, 1).TopBorderColor);
            Assert.AreEqual(RED, GetCellStyle(0, 2).TopBorderColor);
        }
        [Test]
        public void SetBottomBorderColor()
        {
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(2, 0).BottomBorderColor);
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(2, 1).BottomBorderColor);
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(2, 2).BottomBorderColor);
            RegionUtil.SetBottomBorderColor(RED, A1C3, sheet);
            Assert.AreEqual(RED, GetCellStyle(2, 0).BottomBorderColor);
            Assert.AreEqual(RED, GetCellStyle(2, 1).BottomBorderColor);
            Assert.AreEqual(RED, GetCellStyle(2, 2).BottomBorderColor);
        }
        [Test]
        public void SetRightBorderColor()
        {
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(0, 2).RightBorderColor);
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(1, 2).RightBorderColor);
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(2, 2).RightBorderColor);
            RegionUtil.SetRightBorderColor(RED, A1C3, sheet);
            Assert.AreEqual(RED, GetCellStyle(0, 2).RightBorderColor);
            Assert.AreEqual(RED, GetCellStyle(1, 2).RightBorderColor);
            Assert.AreEqual(RED, GetCellStyle(2, 2).RightBorderColor);
        }
        [Test]
        public void SetLeftBorderColor()
        {
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(0, 0).LeftBorderColor);
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(1, 0).LeftBorderColor);
            Assert.AreEqual(DEFAULT_COLOR, GetCellStyle(2, 0).LeftBorderColor);
            RegionUtil.SetLeftBorderColor(RED, A1C3, sheet);
            Assert.AreEqual(RED, GetCellStyle(0, 0).LeftBorderColor);
            Assert.AreEqual(RED, GetCellStyle(1, 0).LeftBorderColor);
            Assert.AreEqual(RED, GetCellStyle(2, 0).LeftBorderColor);
        }

        [Test]
        public void BordersCanBeAddedToNonExistantCells()
        {
            RegionUtil.SetBorderTop(THIN, A1C3, sheet);
            Assert.AreEqual(THIN, GetCellStyle(0, 0).BorderTop);
            Assert.AreEqual(THIN, GetCellStyle(0, 1).BorderTop);
            Assert.AreEqual(THIN, GetCellStyle(0, 2).BorderTop);
        }
        [Test]
        public void BorderColorsCanBeAddedToNonExistantCells()
        {
            RegionUtil.SetTopBorderColor(RED, A1C3, sheet);
            Assert.AreEqual(RED, GetCellStyle(0, 0).TopBorderColor);
            Assert.AreEqual(RED, GetCellStyle(0, 1).TopBorderColor);
            Assert.AreEqual(RED, GetCellStyle(0, 2).TopBorderColor);
        }
    }
}
