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

namespace NPOI.XSSF.usermodel.helpers;

using NPOI.XSSF.usermodel.*;
using NPOI.SS.util.CellRangeAddress;
using NPOI.SS.formula.FormulaParser;
using NPOI.SS.formula.FormulaType;
using NPOI.SS.formula.FormulaRenderer;
using NPOI.SS.usermodel.Row;
using NPOI.SS.usermodel.Cell;
using NPOI.SS.formula.FormulaShifter;
using NPOI.SS.formula.ptg.Ptg;
using NPOI.SS.formula.ptg.AreaPtg;
using NPOI.SS.formula.ptg.AreaErrPtg;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCellFormula;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
using org.Openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;

using java.util.List;
using java.util.ArrayList;

/**
 * @author Yegor Kozlov
 */
public class XSSFRowShifter {
    private XSSFSheet sheet;

    public XSSFRowShifter(XSSFSheet sh) {
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
    public List<CellRangeAddress> ShiftMerged(int startRow, int endRow, int n) {
        List<CellRangeAddress> ShiftedRegions = new ArrayList<CellRangeAddress>();
        //move merged regions completely if they fall within the new region boundaries when they are Shifted
        for (int i = 0; i < sheet.GetNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.GetMergedRegion(i);

            bool inStart = (merged.FirstRow >= startRow || merged.LastRow >= startRow);
            bool inEnd = (merged.FirstRow <= endRow || merged.LastRow <= endRow);

            //don't check if it's not within the Shifted area
            if (!inStart || !inEnd) {
                continue;
            }

            //only shift if the region outside the Shifted rows is not merged too
            if (!ContainsCell(merged, startRow - 1, 0) && !ContainsCell(merged, endRow + 1, 0)) {
                merged.SetFirstRow(merged.GetFirstRow() + n);
                merged.SetLastRow(merged.GetLastRow() + n);
                //have to Remove/add it back
                ShiftedRegions.Add(merged);
                sheet.RemoveMergedRegion(i);
                i = i - 1; // we have to back up now since we Removed one
            }
        }

        //read so it doesn't get Shifted again
        for (CellRangeAddress region : ShiftedRegions) {
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
    private static bool ContainsCell(CellRangeAddress cr, int rowIx, int colIx) {
        if (cr.FirstRow <= rowIx && cr.LastRow >= rowIx
                && cr.FirstColumn <= colIx && cr.LastColumn >= colIx) {
            return true;
        }
        return false;
    }

    /**
     * Updated named ranges
     */
    public void updateNamedRanges(FormulaShifter Shifter) {
        XSSFWorkbook wb = sheet.GetWorkbook();
        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
        for (int i = 0; i < wb.GetNumberOfNames(); i++) {
            XSSFName name = wb.GetNameAt(i);
            String formula = name.GetRefersToFormula();
            int sheetIndex = name.GetSheetIndex();

            Ptg[] ptgs = FormulaParser.Parse(formula, fpb, FormulaType.NAMEDRANGE, sheetIndex);
            if (Shifter.adjustFormula(ptgs, sheetIndex)) {
                String ShiftedFmla = FormulaRenderer.toFormulaString(fpb, ptgs);
                name.SetRefersToFormula(ShiftedFmla);
            }

        }
    }

    /**
     * Update formulas.
     */
    public void updateFormulas(FormulaShifter Shifter) {
        //update formulas on the parent sheet
        updateSheetFormulas(sheet, Shifter);

        //update formulas on other sheets
        XSSFWorkbook wb = sheet.GetWorkbook();
        for (XSSFSheet sh : wb) {
            if (sheet == sh) continue;
            updateSheetFormulas(sh, Shifter);
        }
    }

    private void updateSheetFormulas(XSSFSheet sh, FormulaShifter Shifter) {
        for (Row r : sh) {
            XSSFRow row = (XSSFRow) r;
            updateRowFormulas(row, Shifter);
        }
    }

    private void updateRowFormulas(XSSFRow row, FormulaShifter Shifter) {
        for (Cell c : row) {
            XSSFCell cell = (XSSFCell) c;

            CTCell ctCell = cell.GetCTCell();
            if (ctCell.isSetF()) {
                CTCellFormula f = ctCell.GetF();
                String formula = f.StringValue;
                if (formula.Length > 0) {
                    String ShiftedFormula = ShiftFormula(row, formula, Shifter);
                    if (ShiftedFormula != null) {
                        f.SetStringValue(ShiftedFormula);
                    }
                }

                if (f.isSetRef()) { //Range of cells which the formula applies to.
                    String ref = f.GetRef();
                    String ShiftedRef = ShiftFormula(row, ref, Shifter);
                    if (ShiftedRef != null) f.SetRef(ShiftedRef);
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
    private static String ShiftFormula(XSSFRow row, String formula, FormulaShifter Shifter) {
        XSSFSheet sheet = row.Sheet;
        XSSFWorkbook wb = sheet.GetWorkbook();
        int sheetIndex = wb.GetSheetIndex(sheet);
        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
        Ptg[] ptgs = FormulaParser.Parse(formula, fpb, FormulaType.CELL, sheetIndex);
        String ShiftedFmla = null;
        if (Shifter.adjustFormula(ptgs, sheetIndex)) {
            ShiftedFmla = FormulaRenderer.toFormulaString(fpb, ptgs);
        }
        return ShiftedFmla;
    }

    public void updateConditionalFormatting(FormulaShifter Shifter) {
        XSSFWorkbook wb = sheet.GetWorkbook();
        int sheetIndex = wb.GetSheetIndex(sheet);


        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.Create(wb);
        List<CTConditionalFormatting> cfList = sheet.GetCTWorksheet().GetConditionalFormattingList();
        for(int j = 0; j< cfList.Count; j++){
            CTConditionalFormatting cf = cfList.Get(j);

            ArrayList<CellRangeAddress> cellRanges = new ArrayList<CellRangeAddress>();
            for (Object stRef : cf.GetSqref()) {
                String[] regions = stRef.ToString().split(" ");
                for (int i = 0; i < regions.Length; i++) {
                    cellRanges.Add(CellRangeAddress.ValueOf(regions[i]));
                }
            }

            bool Changed = false;
            List<CellRangeAddress> temp = new ArrayList<CellRangeAddress>();
            for (int i = 0; i < cellRanges.Count; i++) {
                CellRangeAddress craOld = cellRanges.Get(i);
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
                    cfList.Remove(j);
                    continue;
                }
                List<String> refs = new ArrayList<String>();
                for(CellRangeAddress a : temp) refs.Add(a.formatAsString());
                cf.SetSqref(refs);
            }

            for(CTCfRule cfRule : cf.GetCfRuleList()){
                List<String> formulas = cfRule.GetFormulaList();
                for (int i = 0; i < formulas.Count; i++) {
                    String formula = formulas.Get(i);
                    Ptg[] ptgs = FormulaParser.Parse(formula, fpb, FormulaType.CELL, sheetIndex);
                    if (Shifter.adjustFormula(ptgs, sheetIndex)) {
                        String ShiftedFmla = FormulaRenderer.toFormulaString(fpb, ptgs);
                        formulas.Set(i, ShiftedFmla);
                    }
                }
            }
        }
    }

    private static CellRangeAddress ShiftRange(FormulaShifter Shifter, CellRangeAddress cra, int currentExternSheetIx) {
        // FormulaShifter works well in terms of Ptgs - so convert CellRangeAddress to AreaPtg (and back) here
        AreaPtg aptg = new AreaPtg(cra.FirstRow, cra.LastRow, cra.FirstColumn, cra.LastColumn, false, false, false, false);
        Ptg[] ptgs = { aptg, };

        if (!Shifter.adjustFormula(ptgs, currentExternSheetIx)) {
            return cra;
        }
        Ptg ptg0 = ptgs[0];
        if (ptg0 is AreaPtg) {
            AreaPtg bptg = (AreaPtg) ptg0;
            return new CellRangeAddress(bptg.FirstRow, bptg.LastRow, bptg.FirstColumn, bptg.LastColumn);
        }
        if (ptg0 is AreaErrPtg) {
            return null;
        }
        throw new InvalidOperationException("Unexpected Shifted ptg class (" + ptg0.GetClass().GetName() + ")");
    }

}


