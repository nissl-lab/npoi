using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.HSSF.UserModel.Helpers
{
    public class HSSFRowColShifter
    {
        private static POILogger log = POILogFactory.GetLogger(typeof(HSSFRowColShifter));

        /// <summary>
        /// Update formulas.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="formulaShifter"></param>
        public static void UpdateFormulas(ISheet sheet, FormulaShifter formulaShifter)
        {
            //update formulas on the parent sheet
            UpdateSheetFormulas(sheet, formulaShifter);

            //update formulas on other sheets
            IWorkbook wb = sheet.Workbook;
            foreach (ISheet sh in wb)
            {
                if (sheet == sh) continue;
                UpdateSheetFormulas(sh, formulaShifter);
            }
        }
        public static void UpdateSheetFormulas(ISheet sh, FormulaShifter formulashifter)
        {
            foreach (IRow r in sh)
            {
                HSSFRow row = (HSSFRow)r;
                UpdateRowFormulas(row, formulashifter);
            }
        }
        /// <summary>
        /// Update the formulas in specified row using the formula shifting policy specified by shifter
        /// </summary>
        /// <param name="row">the row to update the formulas on</param>
        /// <param name="formulaShifter">the formula shifting policy</param>
        public static void UpdateRowFormulas(IRow row, FormulaShifter formulaShifter)
        {
            ISheet sheet = row.Sheet;
            foreach (ICell c in row)
            {
                HSSFCell cell = (HSSFCell)c;
                String formula = cell.CellFormula;
                if (formula.Length > 0)
                {
                    String shiftedFormula = ShiftFormula(row, formula, formulaShifter);
                    cell.SetCellFormula(shiftedFormula);
                }
            }
        }
        /// <summary>
        /// Shift a formula using the supplied FormulaShifter
        /// </summary>
        /// <param name="row">the row of the cell this formula belongs to. Used to get a reference to the parent workbook.</param>
        /// <param name="formula">the formula to shift</param>
        /// <param name="formulaShifter">the FormulaShifter object that operates on the parsed formula tokens</param>
        /// <returns>the shifted formula if the formula was changed, null if the formula wasn't modified</returns>
        public static String ShiftFormula(IRow row, String formula, FormulaShifter formulaShifter)
        {
            ISheet sheet = row.Sheet;
            IWorkbook wb = sheet.Workbook;
            int sheetIndex = wb.GetSheetIndex(sheet);
            int rowIndex = row.RowNum;
            HSSFEvaluationWorkbook fpb = HSSFEvaluationWorkbook.Create((HSSFWorkbook)wb);

            try
            {
                Ptg[] ptgs = FormulaParser.Parse(formula, fpb, FormulaType.Cell, sheetIndex, rowIndex);
                String shiftedFmla;
                if (formulaShifter.AdjustFormula(ptgs, sheetIndex))
                {
                    shiftedFmla = FormulaRenderer.ToFormulaString(fpb, ptgs);
                }
                else
                {
                    shiftedFmla = formula;
                }
                return shiftedFmla;
            }
            catch (FormulaParseException fpe)
            {
                // Log, but don't change, rather than breaking
                log.Log(POILogger.ERROR, "Error shifting formula on row "+row.RowNum.ToString(),fpe);
                return formula;
            }
        }
    }
}
