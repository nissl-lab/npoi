using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;
using NUnit.Framework.Internal;
using System;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestDays
    {
        //https://support.microsoft.com/en-us/office/days-function-57740535-d549-4395-8728-0f07bff0b9df
        [Test]
        public void TestMicrosoftExample1()
        {
            using(HSSFWorkbook wb = initWorkbook1())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(0).CreateCell(12);
                SS.Util.Utils.AssertDouble(fe, cell, "DAYS(\"15-MAR-2021\",\"1-FEB-2021\")", 42, 0.00000000001);
                SS.Util.Utils.AssertDouble(fe, cell, "DAYS(A2,A3)", 364, 0.00000000001);
            }
        }

        [Test]
        public void TestInvalid()
        {
            using(HSSFWorkbook wb = initWorkbook1())
            {
                HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
                var cell = wb.GetSheetAt(0).GetRow(0).CreateCell(12);
                SS.Util.Utils.AssertError(fe, cell, "DAYS(\"15-XYZ\",\"1-FEB-2021\")", FormulaError.VALUE);
                SS.Util.Utils.AssertError(fe, cell, "DAYS(\"15-MAR-2021\",\"1-XYZ\")", FormulaError.VALUE);
            }
        }

        private HSSFWorkbook initWorkbook1()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, "Data");
            SS.Util.Utils.AddRow(sheet, 1, DateUtil.GetExcelDate(DateTime.Parse("2021-12-31")));
            SS.Util.Utils.AddRow(sheet, 2, DateUtil.GetExcelDate(DateTime.Parse("2021-01-01")));
            return wb;
        }
    }
}
