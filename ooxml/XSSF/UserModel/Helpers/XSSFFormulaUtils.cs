/*
 *  ====================================================================
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for Additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * ====================================================================
 */

namespace NPOI.XSSF.UserModel.Helpers
{

/**
 * Utility to update formulas and named ranges when a sheet name was Changed
 *
 * @author Yegor Kozlov
 */
public class XSSFFormulaUtils {
    private XSSFWorkbook _wb;
    private XSSFEvaluationWorkbook _fpwb;

    public XSSFFormulaUtils(XSSFWorkbook wb) {
        _wb = wb;
        _fpwb = XSSFEvaluationWorkbook.Create(_wb);
    }

    /**
     * Update sheet name in all formulas and named ranges.
     * Called from {@link XSSFWorkbook#SetSheetName(int, String)}
     * <p/>
     * <p>
     * The idea is to parse every formula and render it back to string
     * with the updated sheet name. The FormulaParsingWorkbook passed to the formula Parser
     * is constructed from the old workbook (sheet name is not yet updated) and
     * the FormulaRenderingWorkbook passed to FormulaRenderer#toFormulaString is a custom implementation that
     * returns the new sheet name.
     * </p>
     *
     * @param sheetIndex the 0-based index of the sheet being Changed
     * @param name       the new sheet name
     */
    public void UpdateSheetName(int sheetIndex, String name) {

        /**
         * An instance of FormulaRenderingWorkbook that returns
         */
        FormulaRenderingWorkbook frwb = new FormulaRenderingWorkbook() {

            public ExternalSheet GetExternalSheet(int externSheetIndex) {
                return _fpwb.GetExternalSheet(externSheetIndex);
            }

            public String GetSheetNameByExternSheet(int externSheetIndex) {
                if (externSheetIndex == sheetIndex) return name;
                else return _fpwb.GetSheetNameByExternSheet(externSheetIndex);
            }

            public String ResolveNameXText(NameXPtg nameXPtg) {
                return _fpwb.ResolveNameXText(nameXPtg);
            }

            public String GetNameText(NamePtg namePtg) {
                return _fpwb.GetNameText(namePtg);
            }
        };

        // update named ranges
        for (int i = 0; i < _wb.GetNumberOfNames(); i++) {
            XSSFName nm = _wb.GetNameAt(i);
            if (nm.GetSheetIndex() == -1 || nm.GetSheetIndex() == sheetIndex) {
                updateName(nm, frwb);
            }
        }

        // update formulas
        for (Sheet sh : _wb) {
            for (Row row : sh) {
                for (Cell cell : row) {
                    if (cell.GetCellType() == Cell.CELL_TYPE_FORMULA) {
                        updateFormula((XSSFCell) cell, frwb);
                    }
                }
            }
        }
    }

    /**
     * Parse cell formula and re-assemble it back using the specified FormulaRenderingWorkbook.
     *
     * @param cell the cell to update
     * @param frwb the formula rendering workbbok that returns new sheet name
     */
    private void UpdateFormula(XSSFCell cell, FormulaRenderingWorkbook frwb) {
        CTCellFormula f = cell.GetCTCell().GetF();
        if (f != null) {
            String formula = f.StringValue;
            if (formula != null && formula.Length > 0) {
                int sheetIndex = _wb.GetSheetIndex(cell.GetSheet());
                Ptg[] ptgs = FormulaParser.Parse(formula, _fpwb, FormulaType.CELL, sheetIndex);
                String updatedFormula = FormulaRenderer.toFormulaString(frwb, ptgs);
                if (!formula.Equals(updatedFormula)) f.SetStringValue(updatedFormula);
            }
        }
    }

    /**
     * Parse formula in the named range and re-assemble it  back using the specified FormulaRenderingWorkbook.
     *
     * @param name the name to update
     * @param frwb the formula rendering workbbok that returns new sheet name
     */
    private void UpdateName(XSSFName name, FormulaRenderingWorkbook frwb) {
        String formula = name.GetRefersToFormula();
        if (formula != null) {
            int sheetIndex = name.GetSheetIndex();
            Ptg[] ptgs = FormulaParser.Parse(formula, _fpwb, FormulaType.NAMEDRANGE, sheetIndex);
            String updatedFormula = FormulaRenderer.toFormulaString(frwb, ptgs);
            if (!formula.Equals(updatedFormula)) name.SetRefersToFormula(updatedFormula);
        }
    }
}

}

