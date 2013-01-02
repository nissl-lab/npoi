/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.HSSF.Model
{
    using System;

    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;


    /**
     * HSSF wrapper for the {@link FormulaParser} and {@link FormulaRenderer} 
     * 
     * @author Josh Micich
     */
    public class HSSFFormulaParser
    {

        private static IFormulaParsingWorkbook CreateParsingWorkbook(HSSFWorkbook book)
        {
            return HSSFEvaluationWorkbook.Create(book);
        }

        private HSSFFormulaParser()
        {
            // no instances of this class
        }

        /**
         * Convenience method for parsing cell formulas. see {@link #parse(String, HSSFWorkbook, int)}
         */
        public static Ptg[] Parse(String formula, HSSFWorkbook workbook)
        {
            return FormulaParser.Parse(formula, CreateParsingWorkbook(workbook));
        }

        /**
         * @param formulaType a constant from {@link FormulaType}
         * @return the parsed formula tokens
         */
        public static Ptg[] Parse(String formula, HSSFWorkbook workbook, FormulaType formulaType)
        {
            return FormulaParser.Parse(formula, CreateParsingWorkbook(workbook), formulaType);
        }
        /**
 * @param formula     the formula to parse
 * @param workbook    the parent workbook
 * @param formulaType a constant from {@link FormulaType}
 * @param sheetIndex  the 0-based index of the sheet this formula belongs to.
 * The sheet index is required to resolve sheet-level names. <code>-1</code> means that
 * the scope of the name will be ignored and  the parser will match named ranges only by name
 *
 * @return the parsed formula tokens
 */
        public static Ptg[] Parse(String formula, HSSFWorkbook workbook, FormulaType formulaType, int sheetIndex)
        {
            return FormulaParser.Parse(formula, CreateParsingWorkbook(workbook), formulaType, sheetIndex);
        }


        /**
         * Static method to convert an array of {@link Ptg}s in RPN order
         * to a human readable string format in infix mode.
         * @param book  used for defined names and 3D references
         * @param ptgs  must not be <c>null</c>
         * @return a human readable String
         */
        public static String ToFormulaString(HSSFWorkbook book, Ptg[] ptgs)
        {
            return FormulaRenderer.ToFormulaString(HSSFEvaluationWorkbook.Create(book), ptgs);
        }
    }
}