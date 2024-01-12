using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NUnit.Framework;

namespace TestCases.HSSF.UserModel
{
    [TestFixture]
    public class TestHSSFSheetCopy
    {
        [Test]
        public void TestBasicCopySheet()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            HSSFSheet sheetA = book.CreateSheet("Sheet A") as HSSFSheet;
            sheetA.CreateRow(0).CreateCell(0).SetCellValue("Test case item 1");
            sheetA.CreateRow(1).CreateCell(0).SetCellValue("Test case item 2");
            ISheet sheetB = sheetA.CopySheet("Sheet B", false);
            //Ensure cell values were copied
            Assert.AreEqual(sheetA.GetRow(0).GetCell(0).StringCellValue, sheetB.GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual(sheetA.GetRow(1).GetCell(0).StringCellValue, sheetB.GetRow(1).GetCell(0).StringCellValue);
            //Now test to make sure the copy is independent. Changes to the copy should not affect the original.
            sheetB.GetRow(1).GetCell(0).SetCellValue("This was changed");
            Assert.AreNotEqual(sheetA.GetRow(1).GetCell(0).StringCellValue, sheetB.GetRow(1).GetCell(0).StringCellValue);
        }

        [Test]
        public void TestBasicCopyTo()
        {
            HSSFWorkbook bookA = new HSSFWorkbook();
            HSSFWorkbook bookB = new HSSFWorkbook();
            HSSFWorkbook bookC = new HSSFWorkbook();

            HSSFSheet sheetA = bookA.CreateSheet("Sheet A") as HSSFSheet;
            sheetA.CreateRow(0).CreateCell(0).SetCellValue("Data in the first book");
            HSSFSheet sheetB = bookB.CreateSheet("Sheet B") as HSSFSheet;
            sheetB.CreateRow(0).CreateCell(0).SetCellValue("Data in the second book");
            //Ensure that we can copy into a book that already has a sheet, as well as one that doesn't.
            sheetA.CopyTo(bookB, "Copied Sheet A", false, false);
            sheetA.CopyTo(bookC, "Copied Sheet A", false, false);
            //Ensure the sheet was copied to the 2nd sheet of Book B, not the 1st sheet.
            Assert.AreNotEqual(sheetA.GetRow(0).GetCell(0).StringCellValue, bookB.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual(sheetA.GetRow(0).GetCell(0).StringCellValue, bookB.GetSheetAt(1).GetRow(0).GetCell(0).StringCellValue);
            //Ensure the sheet was copied to the 1st sheet in Book C
            Assert.AreEqual(sheetA.GetRow(0).GetCell(0).StringCellValue, bookC.GetSheetAt(0).GetRow(0).GetCell(0).StringCellValue);
        }

        [Test]
        public void TestColorStyleCopy()
        {
            HSSFWorkbook bookA = new HSSFWorkbook();
            HSSFWorkbook bookB = new HSSFWorkbook();

            bookA.Workbook.CustomPalette.ClearColors();
            bookA.Workbook.CustomPalette.SetColor(0x8, 12, 15, 255); //0x8 is blueish
            bookA.Workbook.CustomPalette.SetColor(0x9, 200, 200, 200); //0x9 is light gray
            bookB.Workbook.CustomPalette.ClearColors();
            bookB.Workbook.CustomPalette.SetColor(0x8, 192, 168, 0); //Throw a color into the destination book so we can see color merge working.

            HSSFSheet sheetA = bookA.CreateSheet("Sheet A") as HSSFSheet;
            ICell cell = sheetA.CreateRow(0).CreateCell(0);
            cell.SetCellValue("I'm a stylish cell!");
            IFont myFont = bookA.CreateFont();
            myFont.FontName = "Times New Roman";
            myFont.IsItalic = true;
            myFont.FontHeightInPoints = 12;
            myFont.Color = 0x8;
            ICellStyle myStyle = bookA.CreateCellStyle();
            myStyle.SetFont(myFont);
            myStyle.FillForegroundColor = 0x9;
            cell.CellStyle = myStyle;

            HSSFSheet sheetB = bookB.CreateSheet("BookB Sheet") as HSSFSheet;
            ICell beeCell = sheetB.CreateRow(0).CreateCell(0);
            ICellStyle styleB = bookB.CreateCellStyle();
            styleB.FillForegroundColor = 0x8;
            beeCell.CellStyle = styleB;
            beeCell.SetCellValue("Hello NPOI");

            //Copy the sheet, and make sure the color, style, and font is correct
            sheetA.CopyTo(bookB, "Copied Sheet A", true, true);
            HSSFSheet theCopy = bookB.GetSheetAt(1) as HSSFSheet;
            ICell copiedCell = theCopy.GetRow(0).GetCell(0);
            //Check that the fill color got copied
            byte[] srcColor = bookA.Workbook.CustomPalette.GetColor(0x9);
            byte[] destColor = bookB.Workbook.CustomPalette.GetColor(copiedCell.CellStyle.FillForegroundColor);
            Assert.IsTrue(srcColor[0]==destColor[0] && srcColor[1]==destColor[1] && srcColor[2]==destColor[2]);
            //Check that the font color got copied
            srcColor = bookA.Workbook.CustomPalette.GetColor(0x8);
            destColor = bookB.Workbook.CustomPalette.GetColor(copiedCell.CellStyle.GetFont(bookB).Color);
            Assert.IsTrue(srcColor[0] == destColor[0] && srcColor[1] == destColor[1] && srcColor[2] == destColor[2]);
            //Check that the fill color of the cell originally in the destination book is still intact
            srcColor = bookB.Workbook.CustomPalette.GetColor(0x8);
            Assert.IsTrue(srcColor[0] == 192 && srcColor[1] == 168 && srcColor[2] == 0);
            //Check that the font made it over okay
            Assert.AreEqual(copiedCell.CellStyle.GetFont(bookB).FontName, myFont.FontName);
        }

        [Test]
        public void TestImageCopy()
        {
            HSSFWorkbook srcBook = HSSFTestDataSamples.OpenSampleWorkbook("Images.xls");
            HSSFWorkbook destBook = new HSSFWorkbook();
            HSSFSheet sheet1 = srcBook.GetSheetAt(0) as HSSFSheet;
            sheet1.CopyTo(destBook, "First Sheet", true, true);

            using (MemoryStream ms = new MemoryStream())
            {
                destBook.Write(ms);
                ms.Position = 0;
                HSSFWorkbook sanityCheck = new HSSFWorkbook(ms);
                //Assert that only one image got copied, because only one image was used on the first page
                Assert.IsTrue(sanityCheck.GetAllPictures().Count == 1);
            }
            HSSFSheet sheet2 = srcBook.GetSheetAt(1) as HSSFSheet;
            sheet2.CopyTo(destBook, "Second Sheet", true, true);
            using (MemoryStream ms = new MemoryStream())
            {
                destBook.Write(ms);
                ms.Position = 0;
                HSSFWorkbook sanityCheck = new HSSFWorkbook(ms);
                //2nd sheet copied, make sure we have two images now, because sheet 2 had one image
                Assert.IsTrue(sanityCheck.GetAllPictures().Count == 2);
            }
        }

        [Test]
        public void CopySheetToWorkbookShouldCopyFormulasOver()
        {
            HSSFWorkbook srcWorkbook = new HSSFWorkbook();
            HSSFSheet srcSheet = srcWorkbook.CreateSheet("Sheet1") as HSSFSheet;

            // Set some values
            IRow row1 = srcSheet.CreateRow((short)0);
            ICell cell = row1.CreateCell((short)0);
            cell.SetCellValue(1);
            ICell cell2 = row1.CreateCell((short)1);
            cell2.SetCellFormula("A1+1");
            HSSFWorkbook destWorkbook = new HSSFWorkbook();
            srcSheet.CopyTo(destWorkbook, srcSheet.SheetName, true, true);

            var destSheet = destWorkbook.GetSheet("Sheet1");
            Assert.NotNull(destSheet);

            Assert.AreEqual(1, destSheet.GetRow(0)?.GetCell(0).NumericCellValue);
            Assert.AreEqual("A1+1", destSheet.GetRow(0)?.GetCell(1).CellFormula);

            destSheet.GetRow(0)?.GetCell(0).SetCellValue(10);
            var evaluator = destWorkbook.GetCreationHelper()
                                        .CreateFormulaEvaluator();

            var destCell = destSheet.GetRow(0)?.GetCell(1);
            evaluator.EvaluateFormulaCell(destCell);
            var destCellValue = evaluator.Evaluate(destCell);

            Assert.AreEqual(11, destCellValue.NumberValue);
        }

        [Test]
        public void CopySheetToWorkbookShouldCopyMergedRegionsOver()
        {
            HSSFWorkbook srcWorkbook = new HSSFWorkbook();
            HSSFSheet srcSheet = srcWorkbook.CreateSheet("Sheet1") as HSSFSheet;

            // Set some merged regions 
            srcSheet.AddMergedRegion(CellRangeAddress.ValueOf("A1:B4"));
            srcSheet.AddMergedRegion(CellRangeAddress.ValueOf("C1:F40"));


            HSSFWorkbook destWorkbook = new HSSFWorkbook();
            srcSheet.CopyTo(destWorkbook, srcSheet.SheetName, true, true);

            var destSheet = destWorkbook.GetSheet("Sheet1");
            Assert.NotNull(destSheet);
            Assert.AreEqual(2, destSheet.MergedRegions.Count);

            Assert.IsTrue(
                new string[]
                {
                    "A1:B4",
                    "C1:F40"
                }
                .SequenceEqual(
                    destSheet.MergedRegions
                             .Select(r => r.FormatAsString())));
        }
    }
}
