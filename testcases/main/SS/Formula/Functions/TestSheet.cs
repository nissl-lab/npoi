using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestSheet
    {
        private static OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 2, 0, 2, null);
        [Test]
        public void TestSheetFunctionWithRealWorkbook()
        {
            var wb = new HSSFWorkbook();
            // Add three sheets: Sheet1, Sheet2, Sheet3
            var sheet1 = wb.CreateSheet("Sheet1");
            var sheet2 = wb.CreateSheet("Sheet2");
            var sheet3 = wb.CreateSheet("Sheet3");

            // Add data
            sheet1.CreateRow(0).CreateCell(0).SetCellValue(123); // A1 in Sheet1
            sheet2.CreateRow(1).CreateCell(0).SetCellValue(456); // A2 in Sheet2

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            // Define formulas and expected results
            String[] formulas = {
                    "SHEET()",
                    "SHEET(A1)",
                    "SHEET(A1:B5)",
                    "SHEET(Sheet2!A2)",
                    "SHEET(\"Sheet3\")",
                    "SHEET(\"invalid\")"
                };

            Object[] expected = {
                    1.0, // current sheet
                    1.0, // A1 in same sheet
                    1.0, // A1:B5 in same sheet
                    2.0, // Sheet2!A2
                    3.0, // Sheet3
                    FormulaError.NA.Code // unknown sheet → #N/A
            };

            // Write formulas to separate cells and evaluate
            var formulaRow = sheet1.CreateRow(1);
            for(int i = 0; i<formulas.Length; i++)
            {
                String formula = formulas[i];
                var cell = formulaRow.CreateCell(i);
                cell.SetCellFormula(formula);
                CellType resultType = fe.EvaluateFormulaCell(cell);

                if(expected[i] is Double)
                {
                    ClassicAssert.AreEqual(CellType.Numeric, resultType,
                            "Unexpected cell type for formula: " + formula);
                    ClassicAssert.AreEqual((Double) expected[i], cell.NumericCellValue,
                                        "Unexpected numeric result for formula: " + formula);
                }
                else if(expected[i] is Byte)
                {
                    ClassicAssert.AreEqual(CellType.Error, resultType,
                            "Unexpected cell type for formula: " + formula);
                    ClassicAssert.AreEqual((byte) expected[i], cell.ErrorCellValue,
                                            "Unexpected error code for formula: " + formula);
                }
            }
        }
    }
}
