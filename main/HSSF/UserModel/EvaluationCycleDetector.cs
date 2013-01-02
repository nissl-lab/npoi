/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Text;
    using System.Collections;

    /**
     * Instances of this class keep track of multiple dependent cell evaluations due
     * to recursive calls to <c>HSSFFormulaEvaluator.internalEvaluate()</c>.
     * The main purpose of this class Is to detect an attempt to evaluate a cell
     * that Is alReady being evaluated. In other words, it detects circular
     * references in spReadsheet formulas.
     * 
     * @author Josh Micich
     */
    class EvaluationCycleDetector
    {

        /**
         * Stores the parameters that identify the evaluation of one cell.<br/>
         */
        private class CellEvaluationFrame
        {

            private HSSFWorkbook _workbook;
            private HSSFSheet _sheet;
            private int _srcRowNum;
            private int _srcColNum;

            public CellEvaluationFrame(HSSFWorkbook workbook, HSSFSheet sheet, int srcRowNum, int srcColNum)
            {
                if (workbook == null)
                {
                    throw new ArgumentException("workbook must not be null");
                }
                if (sheet == null)
                {
                    throw new ArgumentException("sheet must not be null");
                }
                _workbook = workbook;
                _sheet = sheet;
                _srcRowNum = srcRowNum;
                _srcColNum = srcColNum;
            }

            public override bool Equals(Object obj)
            {
                CellEvaluationFrame other = (CellEvaluationFrame)obj;
                if (_workbook != other._workbook)
                {
                    return false;
                }
                if (_sheet != other._sheet)
                {
                    return false;
                }
                if (_srcRowNum != other._srcRowNum)
                {
                    return false;
                }
                if (_srcColNum != other._srcColNum)
                {
                    return false;
                }
                return true;
            }

            public override int GetHashCode ()
            {
                return _workbook.GetHashCode () ^ _sheet.GetHashCode () ^
                    _srcRowNum ^ _srcColNum;
            }

            /**
             * @return human Readable string for debug purposes
             */
            public String FormatAsString()
            {
                return "R=" + _srcRowNum + " C=" + _srcColNum + " ShIx=" + _workbook.GetSheetIndex(_sheet);
            }

            public override String ToString()
            {
                StringBuilder sb = new StringBuilder(64);
                sb.Append(GetType().Name).Append(" [");
                sb.Append(FormatAsString());
                sb.Append("]");
                return sb.ToString();
            }
        }

        private IList _evaluationFrames;

        public EvaluationCycleDetector()
        {
            _evaluationFrames = new ArrayList();
        }

        /**
         * Notifies this evaluation tracker that evaluation of the specified cell Is
         * about to start.<br/>
         * 
         * In the case of a <c>true</c> return code, the caller should
         * continue evaluation of the specified cell, and also be sure to call
         * <c>endEvaluate()</c> when complete.<br/>
         * 
         * In the case of a <c>false</c> return code, the caller should
         * return an evaluation result of
         * <c>ErrorEval.CIRCULAR_REF_ERROR</c>, and not call <c>endEvaluate()</c>.  
         * <br/>
         * @return <c>true</c> if the specified cell has not been visited yet in the current 
         * evaluation. <c>false</c> if the specified cell Is alReady being evaluated.
         */
        public bool StartEvaluate(HSSFWorkbook workbook, HSSFSheet sheet, int srcRowNum, int srcColNum)
        {
            CellEvaluationFrame cef = new CellEvaluationFrame(workbook, sheet, srcRowNum, srcColNum);
            if (_evaluationFrames.Contains(cef))
            {
                return false;
            }
            _evaluationFrames.Add(cef);
            return true;
        }

        /**
         * Notifies this evaluation tracker that the evaluation of the specified
         * cell Is complete. <p/>
         * 
         * Every successful call to <c>startEvaluate</c> must be followed by a
         * call to <c>endEvaluate</c> (recommended in a finally block) to enable
         * proper tracking of which cells are being evaluated at any point in time.<p/>
         * 
         * Assuming a well behaved client, parameters to this method would not be
         * required. However, they have been included to assert correct behaviour,
         * and form more meaningful error messages.
         */
        public void EndEvaluate(HSSFWorkbook workbook, HSSFSheet sheet, int srcRowNum, int srcColNum)
        {
            int nFrames = _evaluationFrames.Count;
            if (nFrames < 1)
            {
                throw new InvalidOperationException("Call to endEvaluate without matching call to startEvaluate");
            }

            nFrames--;
            CellEvaluationFrame cefExpected = (CellEvaluationFrame)_evaluationFrames[nFrames];
            CellEvaluationFrame cefActual = new CellEvaluationFrame(workbook, sheet, srcRowNum, srcColNum);
            if (!cefActual.Equals(cefExpected))
            {
                throw new Exception("Wrong cell specified. "
                        + "Corresponding startEvaluate() call was for cell {"
                        + cefExpected.FormatAsString() + "} this endEvaluate() call Is for cell {"
                        + cefActual.FormatAsString() + "}");
            }
            // else - no problems so pop current frame 
            _evaluationFrames.Remove(nFrames);
        }
    }
}