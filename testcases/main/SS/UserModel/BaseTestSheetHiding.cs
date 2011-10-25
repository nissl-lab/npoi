
namespace TestCases.SS.UserModel
{
    using System;
    using NPOI.HSSF.UserModel;
    using System.IO;
    using TestCases.HSSF;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.UserModel;
    using TestCases.SS;
    /**
     * Tests for how HSSFWorkbook behaves with XLS files
     *  with a WORKBOOK directory entry (instead of the more
     *  usual, Workbook)
     */
    public abstract class BaseTestSheetHiding
    {
        private ITestDataProvider _testDataProvider;
        private String _file1, _file2;

        protected BaseTestSheetHiding(ITestDataProvider testDataProvider, String file1, String file2)
        {
            _testDataProvider = testDataProvider;
            _file1 = file1;
            _file2 = file2;
        }

        private HSSFWorkbook wbH;
        private HSSFWorkbook wbU;

        [TestInitialize]
        public void SetUp()
        {
            wbH = HSSFTestDataSamples.OpenSampleWorkbook(_file1);
            wbU = HSSFTestDataSamples.OpenSampleWorkbook(_file2);
        }

        /**
         * Test that we get the right number of sheets,
         *  with the right text on them, no matter what
         *  the hidden flags are
         */
        [TestMethod]
        public void TestTextSheets()
        {
            // Both should have two sheets
            Assert.AreEqual(2, wbH.NumberOfSheets);
            Assert.AreEqual(2, wbU.NumberOfSheets);

            // All sheets should have one row
            Assert.AreEqual(0, wbH.GetSheetAt(0).LastRowNum);
            Assert.AreEqual(0, wbH.GetSheetAt(1).LastRowNum);
            Assert.AreEqual(0, wbU.GetSheetAt(0).LastRowNum);
            Assert.AreEqual(0, wbU.GetSheetAt(1).LastRowNum);

            // All rows should have one column
            Assert.AreEqual(1, wbH.GetSheetAt(0).GetRow(0).LastCellNum);
            Assert.AreEqual(1, wbH.GetSheetAt(1).GetRow(0).LastCellNum);
            Assert.AreEqual(1, wbU.GetSheetAt(0).GetRow(0).LastCellNum);
            Assert.AreEqual(1, wbU.GetSheetAt(1).GetRow(0).LastCellNum);

            // Text should be sheet based
            Assert.AreEqual("Sheet1A1", wbH.GetSheetAt(0).GetRow(0).GetCell(0).RichStringCellValue.String);
            Assert.AreEqual("Sheet2A1", wbH.GetSheetAt(1).GetRow(0).GetCell(0).RichStringCellValue.String);
            Assert.AreEqual("Sheet1A1", wbU.GetSheetAt(0).GetRow(0).GetCell(0).RichStringCellValue.String);
            Assert.AreEqual("Sheet2A1", wbU.GetSheetAt(1).GetRow(0).GetCell(0).RichStringCellValue.String);
        }

        [TestMethod]
        public void TestSheetHidden()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ISheet sh = wb.CreateSheet("MySheet");

            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));

            wb.SetSheetHidden(0, SheetState.HIDDEN);
            Assert.IsTrue(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));

            wb.SetSheetHidden(0, SheetState.VERY_HIDDEN);
            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsTrue(wb.IsSheetVeryHidden(0));

            wb.SetSheetHidden(0, SheetState.VISIBLE);
            Assert.IsFalse(wb.IsSheetHidden(0));
            Assert.IsFalse(wb.IsSheetVeryHidden(0));

            try
            {
                wb.SetSheetHidden(0, -1);
                Assert.Fail("expectd exception");
            }
            catch (ArgumentException e)
            {
                // ok
            }
            try
            {
                wb.SetSheetHidden(0, 3);
                Assert.Fail("expectd exception");
            }
            catch (ArgumentException e)
            {
                // ok
            }
        }

        /**
         * Check that we can get and set the hidden flags
         *  as expected
         */
        [TestMethod]
        public void TestHideUnHideFlags()
        {
            Assert.IsTrue(wbH.IsSheetHidden(0));
            Assert.IsFalse(wbH.IsSheetHidden(1));
            Assert.IsFalse(wbU.IsSheetHidden(0));
            Assert.IsFalse(wbU.IsSheetHidden(1));
        }

        /**
         * Turn the sheet with none hidden into the one with
         *  one hidden
         */
        [TestMethod]
        public void TestHide()
        {
            wbU.SetSheetHidden(0, SheetState.HIDDEN);
            Assert.IsTrue(wbU.IsSheetHidden(0));
            Assert.IsFalse(wbU.IsSheetHidden(1));
            MemoryStream out1 = new MemoryStream();
            wbU.Write(out1);
            out1.Close();
            HSSFWorkbook wb2 = new HSSFWorkbook(new MemoryStream(out1.ToArray()));
            Assert.IsTrue(wb2.IsSheetHidden(0));
            Assert.IsFalse(wb2.IsSheetHidden(1));
        }

        /**
         * Turn the sheet with one hidden into the one with
         *  none hidden
         */
        [TestMethod]
        public void TestUnHide()
        {
            wbH.SetSheetHidden(0, SheetState.VISIBLE);
            Assert.IsFalse(wbH.IsSheetHidden(0));
            Assert.IsFalse(wbH.IsSheetHidden(1));
            MemoryStream out1 = new MemoryStream();
            wbH.Write(out1);
            out1.Close();
            HSSFWorkbook wb2 = new HSSFWorkbook(new MemoryStream(out1.ToArray()));
            Assert.IsFalse(wb2.IsSheetHidden(0));
            Assert.IsFalse(wb2.IsSheetHidden(1));
        }
    }
}
