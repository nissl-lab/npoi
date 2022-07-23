using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Util
{
    public static class Utils
    {
        public static void AssertString(IFormulaEvaluator fe, ICell cell, string formula, string expectedResult)
        {
            cell.SetCellFormula(formula);
            var result = fe.Evaluate(cell).StringValue;
            fe.ClearAllCachedResultValues();
            Assert.AreEqual(expectedResult, result);
        }
        public static void AssertDouble(IFormulaEvaluator fe, ICell cell, string formula, double expectedResult, double delta=0.0)
        {
            cell.SetCellFormula(formula);
            var result = fe.Evaluate(cell).NumberValue;
            fe.ClearAllCachedResultValues();
            Assert.AreEqual(expectedResult, result, delta);
        }

        public static void AssertError(IFormulaEvaluator fe, ICell cell, string formula, FormulaError expectedError)
        {
            cell.SetCellFormula(formula);
            fe.ClearAllCachedResultValues();

            var result = fe.Evaluate(cell).ErrorValue;
            Assert.AreEqual(expectedError.Code, result);
        }
        public static void AddRow(ISheet sheet, int rowIndex, params object[] values)
        {
            IRow row = sheet.CreateRow(rowIndex);
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i] is string)
                    row.CreateCell(i).SetCellValue((string)values[i]);
                else if (values[i] is null)
                    row.CreateCell(i).SetBlank();
                else if (values[i] is int)
                    row.CreateCell(i).SetCellValue((int)values[i]);
                else if (values[i] == FormulaError.NA)
                    row.CreateCell(i).SetCellFormula("NA()");
                else if (values[i] is bool)
                    row.CreateCell(i).SetCellValue((bool)values[i]);
                else
                    row.CreateCell(i).SetCellValue((double)values[i]);
            }
        }
    }
}
