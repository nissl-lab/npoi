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

namespace NPOI.SS.Formula
{

    /**
     * Used to help optimise cell evaluation result caching by allowing applications to specify which
     * parts of a workbook are <em>final</em>.<br/>
     * The term <b>final</b> is introduced here to denote immutability or 'having constant definition'.
     * This classification refers to potential actions (on the evaluated workbook) by the evaluating
     * application.  It does not refer to operations performed by the evaluator ({@link 
     * WorkbookEvaluator}).<br/>
     * <br/>
     * <b>General guidelines</b>:
     * <ul>
     * <li>a plain value cell can be marked as 'final' if it will not be changed after the first call
     * to {@link WorkbookEvaluator#evaluate(EvaluationCell)}.
     * </li>
     * <li>a formula cell can be marked as 'final' if its formula will not be changed after the first
     * call to {@link WorkbookEvaluator#evaluate(EvaluationCell)}.  This remains true even if changes 
     * in dependent values may cause the evaluated value to change.</li>
     * <li>plain value cells should be marked as 'not final' if their plain value value may change.
     * </li>  
     * <li>formula cells should be marked as 'not final' if their formula definition may change.</li>  
     * <li>cells which may switch between plain value and formula should also be marked as 'not final'.
     * </li>  
     * </ul>
     * <b>Notes</b>:
     * <ul>
     * <li>If none of the spreadsheet cells is expected to have its definition changed after evaluation
     * begins, every cell can be marked as 'final'. This is the most efficient / least resource 
     * intensive option.</li>
     * <li>To retain freedom to change any cell definition at any time, an application may classify all
     * cells as 'not final'.  This freedom comes at the expense of greater memory consumption.</li>
     * <li>For the purpose of these classifications, setting the cached formula result of a cell (for 
     * example in {@link HSSFFormulaEvaluator#evaluateFormulaCell(org.apache.poi.ss.usermodel.Cell)})
     * does not constitute changing the definition of the cell.</li>
     * <li>Updating cells which have been classified as 'final' will cause the evaluator to behave 
     * unpredictably (typically ignoring the update).</li> 
     * </ul>
     * 
     * @author Josh Micich
     */
    public abstract class IStabilityClassifier
    {

        /**
         * Convenience implementation for situations where all cell definitions remain fixed after
         * evaluation begins.
         */
        static IStabilityClassifier TOTALLY_IMMUTABLE = new TotallyImmutable();

        /**
         * Checks if a cell's value(/formula) is fixed - in other words - not expected to be modified
         * between calls to the evaluator. (Note - this is an independent concept from whether a 
         * formula cell's evaluated value may change during successive calls to the evaluator).
         * 
         * @param sheetIndex zero based index into workbook sheet list
         * @param rowIndex zero based row index of cell
         * @param columnIndex zero based column index of cell
         * @return <c>false</c> if the evaluating application may need to modify the specified 
         * cell between calls to the evaluator. 
         */
        public abstract bool IsCellFinal(int sheetIndex, int rowIndex, int columnIndex);
    }

    public class TotallyImmutable : IStabilityClassifier
    {
        public override bool IsCellFinal(int sheetIndex, int rowIndex, int columnIndex)
        {
            return true;
        }
    };
}