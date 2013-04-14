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

using NPOI.SS.Formula;
using System;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.PTG;
namespace NPOI.XSSF.UserModel.Helpers
{
    class XSSFFormulaRenderingWorkbook : IFormulaRenderingWorkbook
    {
        XSSFEvaluationWorkbook _fpwb;
        int _sheetIndex = 0;
        string _name;
        public XSSFFormulaRenderingWorkbook(XSSFEvaluationWorkbook fpwb, int sheetIndex,string name)
        {
            _fpwb = fpwb;
            _sheetIndex = sheetIndex;
            _name = name;
        }
        #region IFormulaRenderingWorkbook Members

        public ExternalSheet GetExternalSheet(int externSheetIndex)
        {
            return _fpwb.GetExternalSheet(externSheetIndex);
        }

        public String GetSheetNameByExternSheet(int externSheetIndex)
        {
            if (externSheetIndex == _sheetIndex) 
                return _name;
            else return _fpwb.GetSheetNameByExternSheet(externSheetIndex);
        }

        public String ResolveNameXText(NameXPtg nameXPtg)
        {
            return _fpwb.ResolveNameXText(nameXPtg);
        }

        public String GetNameText(NamePtg namePtg)
        {
            return _fpwb.GetNameText(namePtg);
        }

        #endregion
    }
    /**
     * Utility to update formulas and named ranges when a sheet name was Changed
     *
     * @author Yegor Kozlov
     */
    public class XSSFFormulaUtils
    {
        private XSSFWorkbook _wb;
        private XSSFEvaluationWorkbook _fpwb;

        public XSSFFormulaUtils(XSSFWorkbook wb)
        {
            _wb = wb;
            _fpwb = XSSFEvaluationWorkbook.Create(_wb);
        }

        /**
         * Update sheet name in all formulas and named ranges.
         * Called from {@link XSSFWorkbook#SetSheetName(int, String)}
         * <p/>
         * <p>
         * The idea is to parse every formula and render it back to string
         * with the updated sheet name. The IFormulaParsingWorkbook passed to the formula Parser
         * is constructed from the old workbook (sheet name is not yet updated) and
         * the FormulaRenderingWorkbook passed to FormulaRenderer#toFormulaString is a custom implementation that
         * returns the new sheet name.
         * </p>
         *
         * @param sheetIndex the 0-based index of the sheet being Changed
         * @param name       the new sheet name
         */
        public void UpdateSheetName(int sheetIndex, String name)
        {

            /**
             * An instance of FormulaRenderingWorkbook that returns
             */
            IFormulaRenderingWorkbook frwb = new XSSFFormulaRenderingWorkbook(_fpwb, sheetIndex, name);
            // update named ranges
            for (int i = 0; i < _wb.NumberOfNames; i++)
            {
                IName nm = _wb.GetNameAt(i);
                if (nm.SheetIndex == -1 || nm.SheetIndex == sheetIndex)
                {
                    UpdateName(nm, frwb);
                }
            }

            // update formulas
            foreach (ISheet sh in _wb)
            {
                foreach (IRow row in sh)
                {
                    foreach (ICell cell in row)
                    {
                        if (cell.CellType == CellType.Formula)
                        {
                            UpdateFormula((XSSFCell)cell, frwb);
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
        private void UpdateFormula(XSSFCell cell, IFormulaRenderingWorkbook frwb)
        {
            CT_CellFormula f = cell.GetCTCell().f;
            if (f != null)
            {
                String formula = f.Value;
                if (formula != null && formula.Length > 0)
                {
                    int sheetIndex = _wb.GetSheetIndex(cell.Sheet);
                    Ptg[] ptgs = FormulaParser.Parse(formula, _fpwb, FormulaType.Cell, sheetIndex);
                    String updatedFormula = FormulaRenderer.ToFormulaString(frwb, ptgs);
                    if (!formula.Equals(updatedFormula)) f.Value = (updatedFormula);
                }
            }
        }

        /**
         * Parse formula in the named range and re-assemble it  back using the specified FormulaRenderingWorkbook.
         *
         * @param name the name to update
         * @param frwb the formula rendering workbbok that returns new sheet name
         */
        private void UpdateName(IName name, IFormulaRenderingWorkbook frwb)
        {
            String formula = name.RefersToFormula;
            if (formula != null)
            {
                int sheetIndex = name.SheetIndex;
                Ptg[] ptgs = FormulaParser.Parse(formula, _fpwb, FormulaType.NamedRange, sheetIndex);
                String updatedFormula = FormulaRenderer.ToFormulaString(frwb, ptgs);
                if (!formula.Equals(updatedFormula)) name.RefersToFormula = (updatedFormula);
            }
        }
    }

}

