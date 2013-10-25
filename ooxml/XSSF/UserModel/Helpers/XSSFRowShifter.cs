/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System.Collections.Generic;
using NPOI.SS.Util;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Formula;
using System;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
namespace NPOI.XSSF.UserModel.Helpers
{
    /**
     * @author Yegor Kozlov
     */
    public class XSSFRowShifter
    {
        private XSSFSheet sheet;

        public XSSFRowShifter(XSSFSheet sh)
        {
            sheet = sh;
        }

        /**
         * Shift merged regions
         *
         * @param startRow the row to start Shifting
         * @param endRow   the row to end Shifting
         * @param n        the number of rows to shift
         * @return an array of affected cell regions
         */
        public List<CellRangeAddress> ShiftMerged(int startRow, int endRow, int n)
        {
            List<CellRangeAddress> ShiftedRegions = new List<CellRangeAddress>();
            //move merged regions completely if they fall within the new region boundaries when they are Shifted
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress merged = sheet.GetMergedRegion(i);

                bool inStart = (merged.FirstRow >= startRow || merged.LastRow >= startRow);
                bool inEnd = (merged.FirstRow <= endRow || merged.LastRow <= endRow);

                //don't check if it's not within the Shifted area
                if (!inStart || !inEnd)
                {
                    continue;
                }

                //only shift if the region outside the Shifted rows is not merged too
                if (!ContainsCell(merged, startRow - 1, 0) && !ContainsCell(merged, endRow + 1, 0))
                {
                    merged.FirstRow = (merged.FirstRow + n);
                    merged.LastRow = (merged.LastRow + n);
                    //have to Remove/add it back
                    ShiftedRegions.Add(merged);
                    sheet.RemoveMergedRegion(i);
                    i = i - 1; // we have to back up now since we Removed one
                }
            }

            //read so it doesn't get Shifted again
            foreach (CellRangeAddress region in ShiftedRegions)
            {
                sheet.AddMergedRegion(region);
            }
            return ShiftedRegions;
        }

        /**
         * Check if the  row and column are in the specified cell range
         *
         * @param cr    the cell range to check in
         * @param rowIx the row to check
         * @param colIx the column to check
         * @return true if the range Contains the cell [rowIx,colIx]
         */
        private static bool ContainsCell(CellRangeAddress cr, int rowIx, int colIx)
        {
            if (cr.FirstRow <= rowIx && cr.LastRow >= rowIx
                    && cr.FirstColumn <= colIx && cr.LastColumn >= colIx)
            {
                return true;
            }
            return false;
        }

        /**
         * Updated named ranges
         */
        public void UpdateNamedRanges(FormulaShifter shifter)
        {
            IWorkbook wb = sheet.Workbook;
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            for (int i = 0; i < wb.NumberOfNames; i++)
            {
                IName name = wb.GetNameAt(i);
                String formula = name.RefersToFormula;
                int sheetIndex = name.SheetIndex;

                Ptg[] ptgs = FormulaParser.Parse(formula, fpb, FormulaType.NamedRange, sheetIndex);
                if (shifter.AdjustFormula(ptgs, sheetIndex))
                {
                    String shiftedFmla = FormulaRenderer.ToFormulaString(fpb, ptgs);
                    name.RefersToFormula = shiftedFmla;
                }

            }
        }

        /**
         * Update formulas.
         */
        public void UpdateFormulas(FormulaShifter shifter)
        {
            //update formulas on the parent sheet
            UpdateSheetFormulas(sheet, shifter);

            //update formulas on other sheets
            IWorkbook wb = sheet.Workbook;
            foreach (XSSFSheet sh in wb)
            {
                if (sheet == sh) continue;
                UpdateSheetFormulas(sh, shifter);
            }
        }

        private void UpdateSheetFormulas(XSSFSheet sh, FormulaShifter Shifter)
        {
            foreach (IRow r in sh)
            {
                XSSFRow row = (XSSFRow)r;
                updateRowFormulas(row, Shifter);
            }
        }

        private void updateRowFormulas(XSSFRow row, FormulaShifter Shifter)
        {
            foreach (ICell c in row)
            {
                XSSFCell cell = (XSSFCell)c;

                CT_Cell ctCell = cell.GetCTCell();
                if (ctCell.IsSetF())
                {
                    CT_CellFormula f = ctCell.f;
                    String formula = f.Value;
                    if (formula.Length > 0)
                    {
                        String ShiftedFormula = ShiftFormula(row, formula, Shifter);
                        if (ShiftedFormula != null)
                        {
                            f.Value = (ShiftedFormula);
                        }
                    }

                    if (f.isSetRef())
                    { //Range of cells which the formula applies to.
                        String ref1 = f.@ref;
                        String ShiftedRef = ShiftFormula(row, ref1, Shifter);
                        if (ShiftedRef != null) f.@ref = ShiftedRef;
                    }
                }

            }
        }

        /**
         * Shift a formula using the supplied FormulaShifter
         *
         * @param row     the row of the cell this formula belongs to. Used to get a reference to the parent workbook.
         * @param formula the formula to shift
         * @param Shifter the FormulaShifter object that operates on the Parsed formula tokens
         * @return the Shifted formula if the formula was Changed,
         *         <code>null</code> if the formula wasn't modified
         */
        private static String ShiftFormula(XSSFRow row, String formula, FormulaShifter Shifter)
        {
            ISheet sheet = row.Sheet;
            IWorkbook wb = sheet.Workbook;
            int sheetIndex = wb.GetSheetIndex(sheet);
            XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
            Ptg[] ptgs = FormulaParser.Parse(formula, fpb, FormulaType.Cell, sheetIndex);
            String ShiftedFmla = null;
            if (Shifter.AdjustFormula(ptgs, sheetIndex))
            {
                ShiftedFmla = FormulaRenderer.ToFormulaString(fpb, ptgs);
            }
            return ShiftedFmla;
        }

        public void UpdateConditionalFormatting(FormulaShifter Shifter) {
        IWorkbook wb = sheet.Workbook;
        int sheetIndex = wb.GetSheetIndex(sheet);


        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
        List<CT_ConditionalFormatting> cfList = sheet.GetCTWorksheet().conditionalFormatting;
        for(int j = 0; cfList!=null&&j< cfList.Count; j++){
            CT_ConditionalFormatting cf = cfList[j];

            List<CellRangeAddress> cellRanges = new List<CellRangeAddress>();
            String[] regions = cf.sqref.ToString().Split(new char[] { ' ' });
            for (int i = 0; i < regions.Length; i++) {
                cellRanges.Add(CellRangeAddress.ValueOf(regions[i]));
            }

            bool Changed = false;
            List<CellRangeAddress> temp = new List<CellRangeAddress>();
            for (int i = 0; i < cellRanges.Count; i++) {
                CellRangeAddress craOld = cellRanges[i];
                CellRangeAddress craNew = ShiftRange(Shifter, craOld, sheetIndex);
                if (craNew == null) {
                    Changed = true;
                    continue;
                }
                temp.Add(craNew);
                if (craNew != craOld) {
                    Changed = true;
                }
            }

            if (Changed) {
                int nRanges = temp.Count;
                if (nRanges == 0) {
                    cfList.RemoveAt(j);
                    continue;
                }
                string refs = string.Empty;
                foreach (CellRangeAddress a in temp)
                {
                    if (refs.Length == 0)
                        refs = a.FormatAsString();
                    else
                        refs += " " + a.FormatAsString();
                }
                cf.sqref = refs;
            }

            foreach(CT_CfRule cfRule in cf.cfRule){
                List<String> formulas = cfRule.formula;
                for (int i = 0; i < formulas.Count; i++) {
                    String formula = formulas[i];
                    Ptg[] ptgs = FormulaParser.Parse(formula, fpb, FormulaType.Cell, sheetIndex);
                    if (Shifter.AdjustFormula(ptgs, sheetIndex)) {
                        String ShiftedFmla = FormulaRenderer.ToFormulaString(fpb, ptgs);
                        formulas[i] =  ShiftedFmla;
                    }
                }
            }
        }
    }

        private static CellRangeAddress ShiftRange(FormulaShifter Shifter, CellRangeAddress cra, int currentExternSheetIx)
        {
            // FormulaShifter works well in terms of Ptgs - so convert CellRangeAddress to AreaPtg (and back) here
            AreaPtg aptg = new AreaPtg(cra.FirstRow, cra.LastRow, cra.FirstColumn, cra.LastColumn, false, false, false, false);
            Ptg[] ptgs = { aptg, };

            if (!Shifter.AdjustFormula(ptgs, currentExternSheetIx))
            {
                return cra;
            }
            Ptg ptg0 = ptgs[0];
            if (ptg0 is AreaPtg)
            {
                AreaPtg bptg = (AreaPtg)ptg0;
                return new CellRangeAddress(bptg.FirstRow, bptg.LastRow, bptg.FirstColumn, bptg.LastColumn);
            }
            if (ptg0 is AreaErrPtg)
            {
                return null;
            }
            throw new InvalidOperationException("Unexpected Shifted ptg class (" + ptg0.GetType().Name + ")");
        }

    }
}


