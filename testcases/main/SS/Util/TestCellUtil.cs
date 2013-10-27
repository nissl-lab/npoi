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

namespace TestCases.SS.Util
{
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;

    using NUnit.Framework;

    /// <summary> The test of class <see cref="CellUtil"/>.
    /// </summary>
    [TestFixture]
    public class TestCellUtil
    {
        /// <summary> Tests the set alignment method with existing style reuse.
        /// </summary>
        [Test]
        public void TestSetAlignment()
        {
            // Create initial cell with default style
            IWorkbook wkb = new HSSFWorkbook();
            ISheet sheet = wkb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);

            // Create a new cell with default style
            ICell cell2 = row.CreateCell(1);

            // Init a new cell style cloned from the one of cell 0
            cell2.CellStyle = wkb.CreateCellStyle();
            cell2.CellStyle.CloneStyleFrom(cell.CellStyle);

            // At this time cell style index should be different
            Assert.AreNotEqual(cell.CellStyle.Index, cell2.CellStyle.Index);

            // Set an arbitraty cell style property to differentiate the two styles
            cell.CellStyle.Alignment = HorizontalAlignment.Right;

            // Try to make the same change so that CellUtil will get existing style
            CellUtil.SetAlignment(cell2, wkb, (short)HorizontalAlignment.Right);

            // Check that cell style has properly been set to HorizontalAlignment.Right
            Assert.AreEqual(cell2.CellStyle.Alignment, HorizontalAlignment.Right);

            // Check that cell style index are the same again
            Assert.AreEqual(cell.CellStyle.Index, cell2.CellStyle.Index);

            // Init a new cell style cloned from the one of cell 0
            cell2.CellStyle = wkb.CreateCellStyle();
            cell2.CellStyle.CloneStyleFrom(cell.CellStyle);

            // Set an arbitraty cell style property to differentiate the two styles
            cell.CellStyle.Alignment = HorizontalAlignment.Left;

            // Try to make different change so that CellUtil will get new style
            CellUtil.SetAlignment(cell2, wkb, (short)HorizontalAlignment.Center);

            // Check that cell style has alignement property set to HorizontalAlignment.Center
            Assert.AreEqual(cell2.CellStyle.Alignment, HorizontalAlignment.Center);

            // Check that cell style index are different
            Assert.AreNotEqual(cell.CellStyle.Index, cell2.CellStyle.Index);
        }
    }
}
