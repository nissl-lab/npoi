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
namespace NPOI.SS.Formula.PTG
{
    /**
     *   Encapsulates logic to convert shared formulaa into non shared equivalent
     */
    public class SharedFormula
    {

        private int _columnWrappingMask;
        private int _rowWrappingMask;

        public SharedFormula(SpreadsheetVersion ssVersion)
        {
            _columnWrappingMask = ssVersion.LastColumnIndex; //"IV" for .xls and  "XFD" for .xlsx
            _rowWrappingMask = ssVersion.LastRowIndex;
        }

        /**
         * Creates a non shared formula from the shared formula counterpart, i.e.
         * Converts the shared formula into the equivalent {@link org.apache.poi.ss.formula.ptg.Ptg} array that it would have,
         * were it not shared.
         *
         * @param ptgs parsed tokens of the shared formula
         * @param formulaRow
         * @param formulaColumn
         */
        public Ptg[] ConvertSharedFormulas(Ptg[] ptgs, int formulaRow, int formulaColumn)
        {

            Ptg[] newPtgStack = new Ptg[ptgs.Length];

            for (int k = 0; k < ptgs.Length; k++)
            {
                Ptg ptg = ptgs[k];
                byte originalOperandClass = unchecked((byte)-1);
                if (!ptg.IsBaseToken)
                {
                    originalOperandClass = ptg.PtgClass;
                }
                if (ptg is RefPtgBase)
                {
                    RefPtgBase refNPtg = (RefPtgBase)ptg;
                    ptg = new RefPtg(FixupRelativeRow(formulaRow, refNPtg.Row, refNPtg.IsRowRelative),
                                         FixupRelativeColumn(formulaColumn, refNPtg.Column, refNPtg.IsColRelative),
                                         refNPtg.IsRowRelative,
                                         refNPtg.IsColRelative);
                    ptg.PtgClass = (originalOperandClass);
                }
                else if (ptg is AreaPtgBase)
                {
                    AreaPtgBase areaNPtg = (AreaPtgBase)ptg;
                    ptg = new AreaPtg(FixupRelativeRow(formulaRow, areaNPtg.FirstRow, areaNPtg.IsFirstRowRelative),
                                    FixupRelativeRow(formulaRow, areaNPtg.LastRow, areaNPtg.IsLastRowRelative),
                                    FixupRelativeColumn(formulaColumn, areaNPtg.FirstColumn, areaNPtg.IsFirstColRelative),
                                    FixupRelativeColumn(formulaColumn, areaNPtg.LastColumn, areaNPtg.IsLastColRelative),
                                    areaNPtg.IsFirstRowRelative,
                                    areaNPtg.IsLastRowRelative,
                                    areaNPtg.IsFirstColRelative,
                                    areaNPtg.IsLastColRelative);
                    ptg.PtgClass = (originalOperandClass);
                }
                else if (ptg is OperandPtg)
                {
                    // Any subclass of OperandPtg is mutable, so it's safest to not share these instances.
                    ptg = ((OperandPtg)ptg).Copy();
                }
                else
                {
                    // all other Ptgs are immutable and can be shared
                }
                newPtgStack[k] = ptg;
            }
            return newPtgStack;
        }

        private int FixupRelativeColumn(int currentcolumn, int column, bool relative)
        {
            if (relative)
            {
                // mask out upper bits to produce 'wrapping' at the maximum column ("IV" for .xls and  "XFD" for .xlsx)
                return (column + currentcolumn) & _columnWrappingMask;
            }
            return column;
        }

        private int FixupRelativeRow(int currentrow, int row, bool relative)
        {
            if (relative)
            {
                return (row + currentrow) & _rowWrappingMask;
            }
            return row;
        }

    }
}