using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Helpers
{
    public abstract class BaseRowColShifter
    {
        //Update named ranges
        public abstract void UpdateNamedRanges(FormulaShifter formulaShifter);

        /**
         * Update formulas.
         */
        public abstract void updateFormulas(FormulaShifter formulaShifter);

        /**
         * Shifts, grows, or shrinks the merged regions due to a row shift
         * ({@link RowShifter}) or column shift ({@link ColumnShifter}).
         * Merged regions that are completely overlaid by shifting will be deleted.
         *
         * @param start the first row or column to be shifted
         * @param end   the last row or column to be shifted
         * @param n     the number of rows or columns to shift
         * @return a list of affected merged regions, excluding contain deleted ones
         */
        public abstract List<CellRangeAddress> shiftMergedRegions(int start, int end, int n);

        /**
         * Update conditional formatting
         * @param formulaShifter The {@link FormulaShifter} to use
         */
        public abstract void updateConditionalFormatting(FormulaShifter formulaShifter);

        /**
         * Shift the Hyperlink anchors (not the hyperlink text, even if the hyperlink
         * is of type LINK_DOCUMENT and refers to a cell that was shifted). Hyperlinks
         * do not track the content they point to.
         *
         * @param formulaShifter the formula shifting policy
         */
        public abstract void updateHyperlinks(FormulaShifter formulaShifter);



        public static CellRangeAddress ShiftRange(FormulaShifter Shifter, CellRangeAddress cra, int currentExternSheetIx)
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
