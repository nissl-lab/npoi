using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCases.SS.Formula.Atp
{
    /// <summary>
    /// Testcase for function XLOOKUP()
    /// </summary>
    [TestFixture]
    public class TestXLookupFunction
    {

        //https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929
        [Test]
        public void TestMicrosoftExample1()
        {
            using(HSSFWorkbook wb = initWorkbook1())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
                SS.Util.Utils.AssertString(fe, cell, "XLOOKUP(F2,B2:B11,D2:D11)", "+55");
                SS.Util.Utils.AssertString(fe, cell, "XLOOKUP(\"Brazil\",B2:B11,D2:D11)", "+55");
                SS.Util.Utils.AssertString(fe, cell, "XLOOKUP(\"brazil\",B2:B11,D2:D11)", "+55");
                //wildcard lookups
                SS.Util.Utils.AssertString(fe, cell, "XLOOKUP(\"brazil\",B2:B11,D2:D11,,2)", "+55");
                SS.Util.Utils.AssertString(fe, cell, "XLOOKUP(\"b*l\",B2:B11,D2:D11,,2)", "+55");
                SS.Util.Utils.AssertString(fe, cell, "XLOOKUP(\"i???a\",B2:B11,D2:D11,,2)", "+91");
            }
        }

        [Test]
        public void TestMicrosoftExample2()
        {
            string formulaText = "XLOOKUP(B2,B5:B14,C5:D14)";
            using(HSSFWorkbook wb = initWorkbook2(8389))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var sheet = wb.GetSheetAt(0);
                var row1 = sheet.GetRow(1);
                string col1 = CellReference.ConvertNumToColString(2);
                string col2 = CellReference.ConvertNumToColString(3);
                string cellRef = $"{col1}:{col2}";
                sheet.SetArrayFormula(formulaText, CellRangeAddress.ValueOf(cellRef));
                fe.EvaluateAll();
                ClassicAssert.AreEqual("Dianne Pugh", row1.GetCell(2).StringCellValue);
                ClassicAssert.AreEqual("Finance", row1.GetCell(3).StringCellValue);
            }
        }

        [Test]
        public void TestMicrosoftExample3()
        {
            using(HSSFWorkbook wb = initWorkbook2(999999))
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
                SS.Util.Utils.AssertError(fe, cell, "XLOOKUP(B2,B5:B14,C5:D14)", FormulaError.NA);

                SS.Util.Utils.AssertString(fe, cell, "XLOOKUP(B2,B5:B14,C5:C14,\"not found\")", "not found");

                string formulaText = "XLOOKUP(B2,B5:B14,C5:D14,\"not found\")";
                var sheet = wb.GetSheetAt(0);
                var row1 = sheet.GetRow(1);
                string col1 = CellReference.ConvertNumToColString(2);
                string col2 = CellReference.ConvertNumToColString(3);
                string cellRef = $"{col1}:{col2}";
                sheet.SetArrayFormula(formulaText, CellRangeAddress.ValueOf(cellRef));
                fe.EvaluateAll();
                ClassicAssert.AreEqual("not found", row1.GetCell(2).StringCellValue);
                ClassicAssert.AreEqual("", row1.GetCell(3).StringCellValue);
            }
        }

        [Test]
        public void TestMicrosoftExample4()
        {
            using(HSSFWorkbook wb = initWorkbook4())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(1).CreateCell(6);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(E2,C2:C7,B2:B7,0,1,1)", 0.24);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(E2,C2:C7,B2:B7,0,1,-1)", 0.24);
            }
        }

        [Test]
        public void TestMicrosoftExample5()
        {
            using(HSSFWorkbook wb = initWorkbook5())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(2).CreateCell(3);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(D2,$B6:$B17,$C6:$C17)", 25000);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(D2,$B6:$B17,XLOOKUP($C3,$C5:$G5,$C6:$G17))", 25000);
            }
        }

        [Test]
        public void TestBinarySearch()
        {
            using(HSSFWorkbook wb = initWorkbook4())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(1).CreateCell(6);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(E2,C2:C7,B2:B7,0,1,2)", 0.24);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(39475,C2:C7,B2:B7,0,0,2)", 0.22);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(39474,C2:C7,B2:B7,0,0,2)", 0);
            }
        }

        [Test]
        public void TestReverseBinarySearch()
        {
            using(HSSFWorkbook wb = initReverseWorkbook4())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(1).CreateCell(6);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(E2,C2:C7,B2:B7,0,1,-2)", 0.24);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(39475,C2:C7,B2:B7,0,0,-2)", 0.22);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(39474,C2:C7,B2:B7,0,0,-2)", 0);
            }
        }

        [Test]
        public void TestReverseBinarySearchWithInvalidValues()
        {
            using(HSSFWorkbook wb = initReverseWorkbook4WithInvalidIncomes())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(1).CreateCell(6);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(E2,C2:C7,B2:B7,0,1,-2)", 0.37);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(9700,C2:C7,B2:B7,0,0,-2)", 0.1);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(39474,C2:C7,B2:B7,0,0,-2)", 0);
            }
        }

        [Test]
        public void TestMicrosoftExample6()
        {
            using(HSSFWorkbook wb = initWorkbook6())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(2).CreateCell(3);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(B3,B6:B10,E6:E10)", 75.28);
                SS.Util.Utils.AssertDouble(fe, cell, "XLOOKUP(C3,B6:B10,E6:E10)", 17.25);
                SS.Util.Utils.AssertDouble(fe, cell, "SUM(XLOOKUP(B3,B6:B10,E6:E10):XLOOKUP(C3,B6:B10,E6:E10))", 110.69);
            }

        }

        private HSSFWorkbook initWorkbook1()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, null, "Counusing", "Abr", "Prefix");
            SS.Util.Utils.AddRow(sheet, 1, null, "China", "CN", "+86", null, "Brazil");
            SS.Util.Utils.AddRow(sheet, 2, null, "India", "IN", "+91");
            SS.Util.Utils.AddRow(sheet, 3, null, "United States", "US", "+1");
            SS.Util.Utils.AddRow(sheet, 4, null, "Indonesia", "ID", "+62");
            SS.Util.Utils.AddRow(sheet, 5, null, "Brazil", "BR", "+55");
            SS.Util.Utils.AddRow(sheet, 6, null, "Pakistan", "PK", "+92");
            SS.Util.Utils.AddRow(sheet, 7, null, "Nigeria", "NG", "+234");
            SS.Util.Utils.AddRow(sheet, 8, null, "Bangladesh", "BD", "+880");
            SS.Util.Utils.AddRow(sheet, 9, null, "Russia", "RU", "+7");
            SS.Util.Utils.AddRow(sheet, 10, null, "Mexico", "MX", "+52");
            return wb;
        }

        private HSSFWorkbook initWorkbook2(int empId)
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, null, "Emp Id", "Employee Name", "Department");
            SS.Util.Utils.AddRow(sheet, 1, null, empId);
            SS.Util.Utils.AddRow(sheet, 3, null, "Emp Id", "Employee Name", "Department");
            SS.Util.Utils.AddRow(sheet, 4, null, 4390, "Ned Lanning", "Marketing");
            SS.Util.Utils.AddRow(sheet, 5, null, 8604, "Margo Hendrix", "Sales");
            SS.Util.Utils.AddRow(sheet, 6, null, 8389, "Dianne Pugh", "Finance");
            SS.Util.Utils.AddRow(sheet, 7, null, 4937, "Earlene McCarty", "Accounting");
            SS.Util.Utils.AddRow(sheet, 8, null, 8299, "Mia Arnold", "Operation");
            SS.Util.Utils.AddRow(sheet, 9, null, 2643, "Jorge Fellows", "Executive");
            SS.Util.Utils.AddRow(sheet, 10, null, 5243, "Rose Winters", "Sales");
            SS.Util.Utils.AddRow(sheet, 11, null, 9693, "Carmela Hahn", "Finance");
            SS.Util.Utils.AddRow(sheet, 12, null, 1636, "Delia Cochran", "Accounting");
            SS.Util.Utils.AddRow(sheet, 13, null, 6703, "Marguerite Cervantes", "Marketing");
            return wb;
        }

        private HSSFWorkbook initWorkbook4()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, null, "Tax Rate", "Max Income", null, "Income", "Tax Rate");
            SS.Util.Utils.AddRow(sheet, 1, null, 0.10, 9700, null, 46523);
            SS.Util.Utils.AddRow(sheet, 2, null, 0.22, 39475);
            SS.Util.Utils.AddRow(sheet, 3, null, 0.24, 84200);
            SS.Util.Utils.AddRow(sheet, 4, null, 0.32, 160726);
            SS.Util.Utils.AddRow(sheet, 5, null, 0.35, 204100);
            SS.Util.Utils.AddRow(sheet, 6, null, 0.37, 510300);
            return wb;
        }

        private HSSFWorkbook initReverseWorkbook4()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, null, "Tax Rate", "Max Income", null, "Income", "Tax Rate");
            SS.Util.Utils.AddRow(sheet, 1, null, 0.37, 510300, null, 46523);
            SS.Util.Utils.AddRow(sheet, 2, null, 0.35, 204100);
            SS.Util.Utils.AddRow(sheet, 3, null, 0.32, 160726);
            SS.Util.Utils.AddRow(sheet, 4, null, 0.24, 84200);
            SS.Util.Utils.AddRow(sheet, 5, null, 0.22, 39475);
            SS.Util.Utils.AddRow(sheet, 6, null, 0.10, 9700);
            return wb;
        }

        private HSSFWorkbook initReverseWorkbook4WithInvalidIncomes()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, null, "Tax Rate", "Max Income", null, "Income", "Tax Rate");
            SS.Util.Utils.AddRow(sheet, 1, null, 0.37, 510300, null, 46523);
            SS.Util.Utils.AddRow(sheet, 2, null, 0.35, "invalid");
            SS.Util.Utils.AddRow(sheet, 3, null, 0.32, "invalid");
            SS.Util.Utils.AddRow(sheet, 4, null, 0.24, "invalid");
            SS.Util.Utils.AddRow(sheet, 5, null, 0.22, "invalid");
            SS.Util.Utils.AddRow(sheet, 6, null, 0.10, 9700);
            return wb;
        }

        private HSSFWorkbook initWorkbook5()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0);
            SS.Util.Utils.AddRow(sheet, 1, null, null, "Quarter", "Gross Profit", "Net Profit", "Profit %");
            SS.Util.Utils.AddRow(sheet, 2, null, null, "Qtr1");
            SS.Util.Utils.AddRow(sheet, 3);
            SS.Util.Utils.AddRow(sheet, 4, null, "Income Statement", "Qtr1", "Qtr2", "Qtr3", "Qtr4", "Total");
            SS.Util.Utils.AddRow(sheet, 5, null, "Total Sales", 50000, 78200, 89500, 91250, 308950);
            SS.Util.Utils.AddRow(sheet, 6, null, "Cost of Sales", -25000, -42050, -59450, -60450, -186950);
            SS.Util.Utils.AddRow(sheet, 7, null, "Gross Profit", 25000, 37150, -30050, -30450, 122000);
            SS.Util.Utils.AddRow(sheet, 8);
            SS.Util.Utils.AddRow(sheet, 9, null, "Depreciation", -899, -791, -202, -412, -2304);
            SS.Util.Utils.AddRow(sheet, 10, null, "Interest", -513, -853, -150, -956, -2472);
            SS.Util.Utils.AddRow(sheet, 11, null, "Earnings before Tax", 23588, 34506, 29698, 29432, 117224);
            SS.Util.Utils.AddRow(sheet, 12);
            SS.Util.Utils.AddRow(sheet, 13, null, "Tax", -4246, -6211, -5346, -5298, -21100);
            SS.Util.Utils.AddRow(sheet, 14);
            SS.Util.Utils.AddRow(sheet, 15, null, "Net Profit", 19342, 28293, 24352, 24134, 96124);
            SS.Util.Utils.AddRow(sheet, 15, null, "Profit %", .293, .278, .234, .236, .269);
            return wb;
        }

        private HSSFWorkbook initWorkbook6()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0);
            SS.Util.Utils.AddRow(sheet, 1, null, "Start", "End", "Total");
            SS.Util.Utils.AddRow(sheet, 2, null, "Grape", "Banana");
            SS.Util.Utils.AddRow(sheet, 3, null, "United States", "US", "+1");
            SS.Util.Utils.AddRow(sheet, 4);
            SS.Util.Utils.AddRow(sheet, 5, null, "Product", "Qty", "Price", "Total");
            SS.Util.Utils.AddRow(sheet, 6, null, "Apple", 23, 0.52, 11.90);
            SS.Util.Utils.AddRow(sheet, 7, null, "Grape", 98, 0.77, 75.28);
            SS.Util.Utils.AddRow(sheet, 8, null, "Pear", 75, 0.24, 18.16);
            SS.Util.Utils.AddRow(sheet, 9, null, "Banana", 95, 0.18, 17.25);
            SS.Util.Utils.AddRow(sheet, 10, null, "Cherry", 42, 0.16, 6.80);
            return wb;
        }

    }
}
