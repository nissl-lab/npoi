/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula
{

    using System;
    using System.Collections;
    using NPOI.SS.Formula.Eval;
    

    /// <summary>
    /// Instances of this class keep track of multiple dependent cell evaluations due
    /// To recursive calls To <see cref="WorkbookEvaluator.Evaluate"/>
    /// The main purpose of this class is To detect an attempt To evaluate a cell
    /// that is already being evaluated. In other words, it detects circular
    /// references in spreadsheet formulas.
    /// </summary>
    /// <remarks>
    /// @author Josh Micich 
    /// </remarks>
    public class EvaluationTracker
    {
        // TODO - consider deleting this class and letting CellEvaluationFrame take care of itself
        private IList _evaluationFrames;
        private IList _currentlyEvaluatingCells;
        private EvaluationCache _cache;

        public EvaluationTracker(EvaluationCache cache)
        {
            _cache = cache;
            _evaluationFrames = new ArrayList();
            _currentlyEvaluatingCells = new ArrayList();
        }

        /**
         * Notifies this evaluation tracker that evaluation of the specified cell Is
         * about To start.<br/>
         *
         * In the case of a <c>true</c> return code, the caller should
         * continue evaluation of the specified cell, and also be sure To call
         * <c>endEvaluate()</c> when complete.<br/>
         *
         * In the case of a <c>null</c> return code, the caller should
         * return an evaluation result of
         * <c>ErrorEval.CIRCULAR_REF_ERROR</c>, and not call <c>endEvaluate()</c>.
         * <br/>
         * @return <c>false</c> if the specified cell is already being evaluated
         */
        public bool StartEvaluate(FormulaCellCacheEntry cce)
        {
            if (cce == null)
            {
                throw new ArgumentException("cellLoc must not be null");
            }
            if (_currentlyEvaluatingCells.Contains(cce))
            {
                return false;
            }
            _currentlyEvaluatingCells.Add(cce);
            _evaluationFrames.Add(new CellEvaluationFrame(cce));
            return true;
        }

        public void UpdateCacheResult(ValueEval result)
        {

            int nFrames = _evaluationFrames.Count;
            if (nFrames < 1)
            {
                throw new InvalidOperationException("Call To endEvaluate without matching call To startEvaluate");
            }
            CellEvaluationFrame frame = (CellEvaluationFrame)_evaluationFrames[nFrames - 1];

            frame.UpdateFormulaResult(result);
        }

        /**
         * Notifies this evaluation tracker that the evaluation of the specified cell is complete. <p/>
         *
         * Every successful call To <c>startEvaluate</c> must be followed by a call To <c>endEvaluate</c> (recommended in a finally block) To enable
         * proper tracking of which cells are being evaluated at any point in time.<p/>
         *
         * Assuming a well behaved client, parameters To this method would not be
         * required. However, they have been included To assert correct behaviour,
         * and form more meaningful error messages.
         */
        public void EndEvaluate(CellCacheEntry cce)
        {

            int nFrames = _evaluationFrames.Count;
            if (nFrames < 1)
            {
                throw new InvalidOperationException("Call To endEvaluate without matching call To startEvaluate");
            }

            nFrames--;
            CellEvaluationFrame frame = (CellEvaluationFrame)_evaluationFrames[nFrames];
            if (cce != frame.GetCCE())
            {
                throw new InvalidOperationException("Wrong cell specified. ");
            }
            // else - no problems so pop current frame
            _evaluationFrames.RemoveAt(nFrames);
            _currentlyEvaluatingCells.Remove(cce);
        }

        public void AcceptFormulaDependency(CellCacheEntry cce)
        {
            // Tell the currently evaluating cell frame that it Has a dependency on the specified
            int prevFrameIndex = _evaluationFrames.Count - 1;
            if (prevFrameIndex < 0)
            {
                // Top level frame, there is no 'cell' above this frame that is using the current cell
            }
            else
            {
                CellEvaluationFrame consumingFrame = (CellEvaluationFrame)_evaluationFrames[prevFrameIndex];
                consumingFrame.AddSensitiveInputCell(cce);
            }
        }

        public void AcceptPlainValueDependency(int bookIndex, int sheetIndex,
                int rowIndex, int columnIndex, ValueEval value)
        {
            // Tell the currently evaluating cell frame that it Has a dependency on the specified
            int prevFrameIndex = _evaluationFrames.Count - 1;
            if (prevFrameIndex < 0)
            {
                // Top level frame, there is no 'cell' above this frame that is using the current cell
            }
            else
            {
                CellEvaluationFrame consumingFrame = (CellEvaluationFrame)_evaluationFrames[prevFrameIndex];
                if (value == BlankEval.instance)
                {
                    consumingFrame.AddUsedBlankCell(bookIndex, sheetIndex, rowIndex, columnIndex);
                }
                else
                {
                    PlainValueCellCacheEntry cce = _cache.GetPlainValueEntry(bookIndex, sheetIndex,
                            rowIndex, columnIndex, value);
                    consumingFrame.AddSensitiveInputCell(cce);
                }
            }
        }
    }
}