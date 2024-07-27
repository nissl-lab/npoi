using NPOI.SS.UserModel;
using NUnit.Framework;
using System;

namespace TestCases.SS.Util
{
    public static class Utils
    {
        public static void AddRow(ISheet sheet, int rowIndex, params object[] values)
        {
            IRow row = sheet.CreateRow(rowIndex);

            for(int i = 0; i < values.Length; i++)
            {
                ICell cell = row.CreateCell(i);

                if(values[i] is null)
                    cell.SetBlank();
                else if(values[i] is string stringValue)
                    cell.SetCellValue(stringValue);
                else if(values[i] is int intValue)
                    cell.SetCellValue(intValue);
                else if(values[i] is FormulaError feValue)
                    cell.SetCellErrorValue(feValue.Code);
                else if(values[i] is bool boolValue)
                    cell.SetCellValue(boolValue);
                else if(values[i] is DateTime dtValue)
                    cell.SetCellValue(dtValue);
                else if(values[i] is double doubleValue)
                    cell.SetCellValue(doubleValue);

#if NET6_0_OR_GREATER
                else if(values[i] is DateOnly doValue)
                    cell.SetCellValue(doValue);
#endif

                else
                    cell.SetCellValue(values[i].ToString());
            }
        }

        public static void AssertString(IFormulaEvaluator fe, ICell cell, string formula, string expectedResult)
        {
            cell.SetCellFormula(formula);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.String, result.CellType);
            Assert.AreEqual(expectedResult, result.StringValue);
        }

        public static void AssertDouble(IFormulaEvaluator fe, ICell cell, string formulaText, double expectedResult, double delta = 0.0)
        {
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Numeric, result.CellType);
            Assert.AreEqual(expectedResult, result.NumberValue, delta);
        }

        public static void AssertBoolean(IFormulaEvaluator fe, ICell cell, string formulaText, bool expectedResult)
        {
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Boolean, result.CellType);
            Assert.AreEqual(expectedResult, result.BooleanValue);
        }

        public static void AssertError(IFormulaEvaluator fe, ICell cell, string formula, FormulaError expectedError)
        {
            cell.SetCellFormula(formula);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Error, result.CellType);
            Assert.AreEqual(expectedError.Code, result.ErrorValue);
        }
    }
}
