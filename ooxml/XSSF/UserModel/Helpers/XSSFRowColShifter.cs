using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Helpers;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.OOXML.XSSF.UserModel.Helpers
{
    public class XSSFRowColShifter
    {
        /// <summary>
        /// Updated named ranges
        /// </summary>
        /// <param name="shifter"></param>
        public static void UpdateNamedRanges(ISheet sheet, FormulaShifter shifter)
        {
            IWorkbook wb = sheet.Workbook;
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            foreach (IName name in wb.GetAllNames())
            {
                string formula = name.RefersToFormula;
                int sheetIndex = name.SheetIndex;

                SS.Formula.PTG.Ptg[] ptgs = 
                    FormulaParser.Parse(formula, fpb, FormulaType.NamedRange, sheetIndex, -1);

                if (shifter.AdjustFormula(ptgs, sheetIndex))
                {
                    string shiftedFmla = FormulaRenderer.ToFormulaString(fpb, ptgs);
                    name.RefersToFormula = shiftedFmla;
                }
            }
        }

        /// <summary>
        /// Update formulas.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="shifter"></param>
        public static void UpdateFormulas(ISheet sheet, FormulaShifter shifter)
        {
            //update formulas on the parent sheet
            UpdateSheetFormulas(sheet, shifter);

            //update formulas on other sheets
            IWorkbook wb = sheet.Workbook;
            foreach (XSSFSheet sh in wb)
            {
                if (sheet == sh)
                {
                    continue;
                }

                UpdateSheetFormulas(sh, shifter);
            }
        }

        public static void UpdateSheetFormulas(ISheet sh, FormulaShifter Shifter)
        {
            foreach (IRow r in sh)
            {
                XSSFRow row = (XSSFRow)r;
                UpdateRowFormulas(row, Shifter);
            }
        }

        /// <summary>
        /// Update the formulas in specified row using the formula shifting 
        /// policy specified by shifter
        /// </summary>
        /// <param name="row">the row to update the formulas on</param>
        /// <param name="Shifter">the formula shifting policy</param>
        public static void UpdateRowFormulas(IRow row, FormulaShifter Shifter)
        {
            XSSFSheet sheet = (XSSFSheet)row.Sheet;
            foreach (ICell c in row)
            {
                XSSFCell cell = (XSSFCell)c;

                CT_Cell ctCell = cell.GetCTCell();
                if (ctCell.IsSetF())
                {
                    CT_CellFormula f = ctCell.f;
                    string formula = f.Value;
                    if (formula.Length > 0)
                    {
                        string ShiftedFormula = ShiftFormula(row, formula, Shifter);
                        if (ShiftedFormula != null)
                        {
                            f.Value = ShiftedFormula;
                            if (f.t == ST_CellFormulaType.shared)
                            {
                                int si = (int)f.si;
                                CT_CellFormula sf = sheet.GetSharedFormula(si);
                                sf.Value = ShiftedFormula;
                            }
                        }
                    }

                    if (f.isSetRef())
                    { //Range of cells which the formula applies to.
                        string ref1 = f.@ref;
                        string ShiftedRef = ShiftFormula(row, ref1, Shifter);
                        if (ShiftedRef != null)
                        {
                            f.@ref = ShiftedRef;
                    }
                }
            }
        }
        }

        /// <summary>
        /// Shift a formula using the supplied FormulaShifter
        /// </summary>
        /// <param name="row"> the row of the cell this formula belongs to. 
        /// Used to get a reference to the parent workbook.</param>
        /// <param name="formula">the formula to shift</param>
        /// <param name="Shifter">the FormulaShifter object that operates on 
        /// the Parsed formula tokens</param>
        /// <returns>the Shifted formula if the formula was changed, null  if 
        /// the formula wasn't modified</returns>
        private static string ShiftFormula(IRow row, string formula, FormulaShifter Shifter)
        {
            ISheet sheet = row.Sheet;
            IWorkbook wb = sheet.Workbook;
            int sheetIndex = wb.GetSheetIndex(sheet);
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            try
            {
                SS.Formula.PTG.Ptg[] ptgs = 
                    FormulaParser.Parse(formula, fpb, FormulaType.Cell, sheetIndex, -1);
                string ShiftedFmla = null;
                if (Shifter.AdjustFormula(ptgs, sheetIndex))
                {
                    ShiftedFmla = FormulaRenderer.ToFormulaString(fpb, ptgs);
                }

                return ShiftedFmla;
            }
            catch (FormulaParseException fpe)
            {
                // Log, but don't change, rather than breaking
                Console.WriteLine($"Error shifting formula on row {row.RowNum}, {fpe}");
                return formula;
            }
        }

        /// <summary>
        /// Shift the Hyperlink anchors(not the hyperlink text, even if the hyperlink is of type LINK_DOCUMENT and refers to a cell that was shifted). Hyperlinks  do not track the content they point to.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="shifter"></param>
        public static void UpdateHyperlinks(ISheet sheet, FormulaShifter shifter)
        {
            XSSFSheet xsheet = (XSSFSheet)sheet;
            int sheetIndex = xsheet.GetWorkbook().GetSheetIndex(sheet);
            List<IHyperlink> hyperlinkList = sheet.GetHyperlinkList();

            foreach (IHyperlink hyperlink1 in hyperlinkList)
            {
                XSSFHyperlink hyperlink = hyperlink1 as XSSFHyperlink;
                string cellRef = hyperlink.CellRef;
                CellRangeAddress cra = CellRangeAddress.ValueOf(cellRef);
                CellRangeAddress shiftedRange = 
                    BaseRowColShifter.ShiftRange(shifter, cra, sheetIndex);

                if (shiftedRange != null && shiftedRange != cra)
                {
                    // shiftedRange should not be null. If shiftedRange is null, that means
                    // that a hyperlink wasn't deleted at the beginning of shiftRows when
                    // identifying rows that should be removed because they will be overwritten
                    hyperlink.SetCellReference(shiftedRange.FormatAsString());
                }
            }
        }

        public static void UpdateConditionalFormatting(ISheet sheet, FormulaShifter Shifter)
        {
            XSSFSheet xsheet = (XSSFSheet)sheet;
            XSSFWorkbook wb = xsheet.Workbook as XSSFWorkbook;
            int sheetIndex = wb.GetSheetIndex(sheet);
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            CT_Worksheet ctWorksheet = xsheet.GetCTWorksheet();
            List<CT_ConditionalFormatting> conditionalFormattingArray = 
                ctWorksheet.conditionalFormatting;

            // iterate backwards due to possible calls to ctWorksheet.removeConditionalFormatting(j)
            for (int j = conditionalFormattingArray.Count - 1; j >= 0; j--)
            {
                CT_ConditionalFormatting cf = conditionalFormattingArray[j];
                List<CellRangeAddress> cellRanges = new List<CellRangeAddress>();
                string[] regions = cf.sqref.ToString().Split(new char[] { ' ' });

                for (int i = 0; i < regions.Length; i++)
                {
                    cellRanges.Add(CellRangeAddress.ValueOf(regions[i]));
                }

                bool Changed = false;
                List<CellRangeAddress> temp = new List<CellRangeAddress>();
                for (int i = 0; i < cellRanges.Count; i++)
                {
                    CellRangeAddress craOld = cellRanges[i];
                    CellRangeAddress craNew = 
                        BaseRowColShifter.ShiftRange(Shifter, craOld, sheetIndex);

                    if (craNew == null)
                    {
                        Changed = true;
                        continue;
                    }

                    temp.Add(craNew);
                    if (craNew != craOld)
                    {
                        Changed = true;
                    }
                }

                if (Changed)
                {
                    int nRanges = temp.Count;
                    if (nRanges == 0)
                    {
                        conditionalFormattingArray.RemoveAt(j);
                        continue;
                    }

                    string refs = string.Empty;
                    foreach (CellRangeAddress a in temp)
                    {
                        if (refs.Length == 0)
                        {
                            refs = a.FormatAsString();
                        }
                        else
                        {
                            refs += " " + a.FormatAsString();
                    }
                    }

                    cf.sqref = refs;
                }

                foreach (CT_CfRule cfRule in cf.cfRule)
                {
                    List<string> formulas = cfRule.formula;
                    for (int i = 0; i < formulas.Count; i++)
                    {
                        string formula = formulas[i];
                        SS.Formula.PTG.Ptg[] ptgs = 
                            FormulaParser.Parse(formula, fpb, FormulaType.Cell, sheetIndex, -1);

                        if (Shifter.AdjustFormula(ptgs, sheetIndex))
                        {
                            string ShiftedFmla = FormulaRenderer.ToFormulaString(fpb, ptgs);
                            formulas[i] = ShiftedFmla;
                        }
                    }
                }
            }
        }
    }
}
