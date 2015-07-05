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

using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;
namespace NPOI.XSSF.UserModel
{

    /**
     * Test asserts the POI produces &lt;cols&gt; element that could be read and properly interpreted by the MS Excel.
     * For specification of the "cols" element see the chapter 3.3.1.16 of the "Office Open XML Part 4 - Markup Language Reference.pdf".
     * The specification can be downloaded at http://www.ecma-international.org/publications/files/ECMA-ST/Office%20Open%20XML%201st%20edition%20Part%204%20(PDF).zip.
     * 
     * <p><em>
     * The Test saves xlsx file on a disk if the system property is Set:
     * -Dpoi.test.XSSF.output.dir=${workspace_loc}/poi/build/xssf-output
     * </em>
     * 
     */
    [TestFixture]
    public class TestXSSFColGrouping
    {

        /**
         * Tests that POI doesn't produce "col" elements without "width" attribute. 
         * POI-52186
         */
        [Test]
        public void TestNoColsWithoutWidthWhenGrouping()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("test");

            sheet.SetColumnWidth(4, 5000);
            sheet.SetColumnWidth(5, 5000);

            sheet.GroupColumn((short)4, (short)7);
            sheet.GroupColumn((short)9, (short)12);

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb, "testNoColsWithoutWidthWhenGrouping");
            sheet = (XSSFSheet)wb.GetSheet("test");

            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            //logger.log(POILogger.DEBUG, "test52186/cols:" + cols);
            foreach (CT_Col col in cols.GetColList())
            {
                Assert.IsTrue(col.IsSetWidth(), "Col width attribute is unset: " + col.ToString());
            }
        }

        /**
         * Tests that POI doesn't produce "col" elements without "width" attribute. 
         * POI-52186
         */
        [Test]
        public void TestNoColsWithoutWidthWhenGroupingAndCollapsing()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("test");

            sheet.SetColumnWidth(4, 5000);
            sheet.SetColumnWidth(5, 5000);

            sheet.GroupColumn((short)4, (short)5);

            sheet.SetColumnGroupCollapsed(4, true);

            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            //logger.log(POILogger.DEBUG, "test52186_2/cols:" + cols);

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb, "testNoColsWithoutWidthWhenGroupingAndCollapsing");
            sheet = (XSSFSheet)wb.GetSheet("test");

            for (int i = 4; i <= 5; i++)
            {
                Assert.AreEqual(5000, sheet.GetColumnWidth(i), "Unexpected width of column " + i);
            }
            cols = sheet.GetCTWorksheet().GetColsArray(0);
            foreach (CT_Col col in cols.GetColList())
            {
                Assert.IsTrue(col.IsSetWidth(), "Col width attribute is unset: " + col.ToString());
            }
        }

        /**
         * Test the cols element is correct in case of NumericRanges.OVERLAPS_2_WRAPS
         */
        [Test]
        public void TestMergingOverlappingCols_OVERLAPS_2_WRAPS()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("test");

            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            CT_Col col = cols.AddNewCol();
            col.min=(1 + 1);
            col.max=(4 + 1);
            col.width=(20);
            col.customWidth=(true);

            sheet.GroupColumn((short)2, (short)3);

            sheet.GetCTWorksheet().GetColsArray(0);
            //logger.log(POILogger.DEBUG, "testMergingOverlappingCols_OVERLAPS_2_WRAPS/cols:" + cols);

            Assert.AreEqual(0, cols.GetColArray(0).outlineLevel);
            Assert.AreEqual(2, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(2, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(true, cols.GetColArray(0).customWidth);

            Assert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            Assert.AreEqual(3, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(4, cols.GetColArray(1).max); // 1 based        
            Assert.AreEqual(true, cols.GetColArray(1).customWidth);

            Assert.AreEqual(0, cols.GetColArray(2).outlineLevel);
            Assert.AreEqual(5, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(5, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(true, cols.GetColArray(2).customWidth);

            Assert.AreEqual(3, cols.sizeOfColArray());

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb, "testMergingOverlappingCols_OVERLAPS_2_WRAPS");
            sheet = (XSSFSheet)wb.GetSheet("test");

            for (int i = 1; i <= 4; i++)
            {
                Assert.AreEqual( 20 * 256, sheet.GetColumnWidth(i), "Unexpected width of column " + i);
            }
        }

        /**
         * Test the cols element is correct in case of NumericRanges.OVERLAPS_1_WRAPS
         */
        [Test]
        public void TestMergingOverlappingCols_OVERLAPS_1_WRAPS()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("test");

            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            CT_Col col = cols.AddNewCol();
            col.min=(2 + 1);
            col.max=(4 + 1);
            col.width=(20);
            col.customWidth=(true);

            sheet.GroupColumn((short)1, (short)5);

            cols = sheet.GetCTWorksheet().GetColsArray(0);
            //logger.log(POILogger.DEBUG, "testMergingOverlappingCols_OVERLAPS_1_WRAPS/cols:" + cols);

            Assert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            Assert.AreEqual(2, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(2, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(0).customWidth);

            Assert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            Assert.AreEqual(3, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(5, cols.GetColArray(1).max); // 1 based        
            Assert.AreEqual(true, cols.GetColArray(1).customWidth);

            Assert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            Assert.AreEqual(6, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(6, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(2).customWidth);

            Assert.AreEqual(3, cols.sizeOfColArray());

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb, "testMergingOverlappingCols_OVERLAPS_1_WRAPS");
            sheet = (XSSFSheet)wb.GetSheet("test");

            for (int i = 2; i <= 4; i++)
            {
                Assert.AreEqual(20 * 256, sheet.GetColumnWidth(i), "Unexpected width of column " + i);
            }
        }

        /**
         * Test the cols element is correct in case of NumericRanges.OVERLAPS_1_MINOR
         */
        [Test]
        public void TestMergingOverlappingCols_OVERLAPS_1_MINOR()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("test");

            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            CT_Col col = cols.AddNewCol();
            col.min=(2 + 1);
            col.max=(4 + 1);
            col.width=(20);
            col.customWidth=(true);

            sheet.GroupColumn((short)3, (short)5);

            cols = sheet.GetCTWorksheet().GetColsArray(0);
            //logger.log(POILogger.DEBUG, "testMergingOverlappingCols_OVERLAPS_1_MINOR/cols:" + cols);

            Assert.AreEqual(0, cols.GetColArray(0).outlineLevel);
            Assert.AreEqual(3, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(3, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(true, cols.GetColArray(0).customWidth);

            Assert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            Assert.AreEqual(4, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(5, cols.GetColArray(1).max); // 1 based        
            Assert.AreEqual(true, cols.GetColArray(1).customWidth);

            Assert.AreEqual(1, cols.GetColArray(2).outlineLevel);
            Assert.AreEqual(6, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(6, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(2).customWidth);

            Assert.AreEqual(3, cols.sizeOfColArray());

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb, "testMergingOverlappingCols_OVERLAPS_1_MINOR");
            sheet = (XSSFSheet)wb.GetSheet("test");

            for (int i = 2; i <= 4; i++)
            {
                Assert.AreEqual( 20 * 256, sheet.GetColumnWidth(i), "Unexpected width of column " + i);
            }
            Assert.AreEqual( sheet.DefaultColumnWidth * 256, sheet.GetColumnWidth(5), "Unexpected width of column " + 5);
        }

        /**
         * Test the cols element is correct in case of NumericRanges.OVERLAPS_2_MINOR
         */
        [Test]
        public void TestMergingOverlappingCols_OVERLAPS_2_MINOR()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("test");

            CT_Cols cols = sheet.GetCTWorksheet().GetColsArray(0);
            CT_Col col = cols.AddNewCol();
            col.min=(2 + 1);
            col.max=(4 + 1);
            col.width=(20);
            col.customWidth=(true);

            sheet.GroupColumn((short)1, (short)3);

            cols = sheet.GetCTWorksheet().GetColsArray(0);
            //logger.log(POILogger.DEBUG, "testMergingOverlappingCols_OVERLAPS_2_MINOR/cols:" + cols);

            Assert.AreEqual(1, cols.GetColArray(0).outlineLevel);
            Assert.AreEqual(2, cols.GetColArray(0).min); // 1 based
            Assert.AreEqual(2, cols.GetColArray(0).max); // 1 based
            Assert.AreEqual(false, cols.GetColArray(0).customWidth);

            Assert.AreEqual(1, cols.GetColArray(1).outlineLevel);
            Assert.AreEqual(3, cols.GetColArray(1).min); // 1 based
            Assert.AreEqual(4, cols.GetColArray(1).max); // 1 based        
            Assert.AreEqual(true, cols.GetColArray(1).customWidth);

            Assert.AreEqual(0, cols.GetColArray(2).outlineLevel);
            Assert.AreEqual(5, cols.GetColArray(2).min); // 1 based
            Assert.AreEqual(5, cols.GetColArray(2).max); // 1 based
            Assert.AreEqual(true, cols.GetColArray(2).customWidth);

            Assert.AreEqual(3, cols.sizeOfColArray());

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb, "testMergingOverlappingCols_OVERLAPS_2_MINOR");
            sheet = (XSSFSheet)wb.GetSheet("test");

            for (int i = 2; i <= 4; i++)
            {
                Assert.AreEqual(20 * 256, sheet.GetColumnWidth(i), "Unexpected width of column " + i);
            }
            Assert.AreEqual(sheet.DefaultColumnWidth * 256, sheet.GetColumnWidth(1),"Unexpected width of column " + 1 );
        }

    }

}