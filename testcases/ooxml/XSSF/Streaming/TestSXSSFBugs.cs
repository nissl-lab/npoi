using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System.IO;
using TestCases.SS.UserModel;

namespace TestCases.XSSF.Streaming
{
    [TestFixture]
    public class TestSXSSFBugs : BaseTestBugzillaIssues
    {
        public TestSXSSFBugs()
            : base(SXSSFITestDataProvider.instance)
        {

        }
        // override some tests which do not work for SXSSF
        [Ignore("cloneSheet() not implemented")]  public override void Bug18800() { /* cloneSheet() not implemented */ }
        [Ignore("cloneSheet() not implemented")]  public override void Bug22720() { /* cloneSheet() not implemented */ }
        [Ignore("Evaluation is not fully supported")]  public override void Bug47815() { /* Evaluation is not supported */ }
        [Ignore("Evaluation is not fully supported")]  public override void Test58113() { /* Evaluation is not supported */ }
        [Ignore("Evaluation is not fully supported")] public override void Bug46729_testMaxFunctionArguments() { /* Evaluation is not supported */ }
        [Ignore("Reading data is not supported")] public override void Bug57798() { /* Reading data is not supported */ }

        [Test]
        public void Bug49253()
        {
            IWorkbook wb1 = new SXSSFWorkbook();
            IWorkbook wb2 = new SXSSFWorkbook();
            CellRangeAddress cra = CellRangeAddress.ValueOf("C2:D3");

            // No print settings before repeating
            ISheet s1 = wb1.CreateSheet();
            s1.RepeatingColumns = (cra);
            s1.RepeatingRows = (cra);

            IPrintSetup ps1 = s1.PrintSetup;
            Assert.AreEqual(false, ps1.ValidSettings);
            Assert.AreEqual(false, ps1.Landscape);


            // Had valid print settings before repeating
            ISheet s2 = wb2.CreateSheet();
            IPrintSetup ps2 = s2.PrintSetup;

            ps2.Landscape = (false);
            Assert.AreEqual(true, ps2.ValidSettings);
            Assert.AreEqual(false, ps2.Landscape);
            s2.RepeatingColumns = (cra);
            s2.RepeatingRows = (cra);

            ps2 = s2.PrintSetup;
            Assert.AreEqual(true, ps2.ValidSettings);
            Assert.AreEqual(false, ps2.Landscape);

            wb1.Close();
            wb2.Close();
        }

        // bug 60197: setSheetOrder should update sheet-scoped named ranges to maintain references to the sheets before the re-order
        [Test]
        override
        public void bug60197_NamedRangesReferToCorrectSheetWhenSheetOrderIsChanged()
        {
            try
            {
                base.bug60197_NamedRangesReferToCorrectSheetWhenSheetOrderIsChanged();
            }
            catch (RuntimeException e)
            {
                var cause = e.InnerException;
                if (cause is IOException && cause.Message == "Stream closed")
                {
                    // expected on the second time that _testDataProvider.writeOutAndReadBack(SXSSFWorkbook) is called
                    // if the test makes it this far, then we know that XSSFName sheet indices are updated when sheet
                    // order is changed, which is the purpose of this test. Therefore, consider this a passing test.
                }
                else
                {
                    throw e;
                }
            }
        }

        /**
         * Default Column style
         */
        [Test]
        public void Test51037()
        {
            SXSSFWorkbook wb = new SXSSFWorkbook();

            ICellStyle blueStyle = wb.CreateCellStyle();
            blueStyle.FillForegroundColor = IndexedColors.Aqua.Index;
            blueStyle.FillPattern = FillPattern.SolidForeground;

            ICellStyle pinkStyle = wb.CreateCellStyle();
            pinkStyle.FillForegroundColor = IndexedColors.Pink.Index;
            pinkStyle.FillPattern = FillPattern.SolidForeground;

            SXSSFSheet s1 = wb.CreateSheet("Pretty columns") as SXSSFSheet;

            s1.SetDefaultColumnStyle(4, blueStyle);
            s1.SetDefaultColumnStyle(6, pinkStyle);

            SXSSFRow r3 = s1.CreateRow(3) as SXSSFRow;
            r3.CreateCell(0).SetCellValue("The");
            r3.CreateCell(1).SetCellValue("quick");
            r3.CreateCell(2).SetCellValue("brown");
            r3.CreateCell(3).SetCellValue("fox");
            r3.CreateCell(4).SetCellValue("jumps");
            r3.CreateCell(5).SetCellValue("over");
            r3.CreateCell(6).SetCellValue("the");
            r3.CreateCell(7).SetCellValue("lazy");
            r3.CreateCell(8).SetCellValue("dog");

            Assert.AreEqual(blueStyle.Index, r3.GetCell(4).CellStyle.Index);
            Assert.AreEqual(pinkStyle.Index, r3.GetCell(6).CellStyle.Index);

            // Save, re-load and re-check 
            var wb2 = (XSSFWorkbook)_testDataProvider.WriteOutAndReadBack(wb);
            XSSFSheet wb2Sheet= wb2.GetSheetAt(0) as XSSFSheet;
            XSSFRow wb2R3 = wb2Sheet.GetRow(3) as XSSFRow;
            Assert.AreEqual(blueStyle.Index, wb2R3.GetCell(4).CellStyle.Index);
            Assert.AreEqual(pinkStyle.Index, wb2R3.GetCell(6).CellStyle.Index);

            wb.Close();
        }
    }
}
