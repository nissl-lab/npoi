using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public interface ArrayFunction
    {
        /// <summary>
        /// - Excel uses the error code #NUM! instead of IEEE NaN, so when numeric functions evaluate to Double#NaN be sure to translate the result to ErrorEval#NUM_ERROR.
        /// </summary>
        /// <param name="args">the evaluated function arguments.  Empty values are represented with BlankEval or MissingArgEval, never <code>null</code></param>
        /// <param name="srcRowIndex">row index of the cell containing the formula under evaluation</param>
        /// <param name="srcColumnIndex">column index of the cell containing the formula under evaluation</param>
        /// <returns> The evaluated result, possibly an ErrorEval, never <code>null</code></returns>
        ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex);
    }
}
