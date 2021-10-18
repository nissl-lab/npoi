using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{

    /**
     * @author Robert Hulbert
     * Common Interface for any excel built-in function that has implemented array formula functionality.
     */
    public interface ArrayFunction
    {
        /**
     * @param args the evaluated function arguments.  Empty values are represented with
     * {@link BlankEval} or {@link MissingArgEval}, never <code>null</code>.
     * @param srcRowIndex row index of the cell containing the formula under evaluation
     * @param srcColumnIndex column index of the cell containing the formula under evaluation
     * @return The evaluated result, possibly an {@link ErrorEval}, never <code>null</code>.
     * <b>Note</b> - Excel uses the error code <i>#NUM!</i> instead of IEEE <i>NaN</i>, so when
     * numeric functions evaluate to {@link Double#NaN} be sure to translate the result to {@link
     * ErrorEval#NUM_ERROR}.
     */

        ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex);

        /**
         * Evaluate an array function with two arguments.
         *
         * @param arg0 the first function argument. Empty values are represented with
         *        {@link BlankEval} or {@link MissingArgEval}, never <code>null</code>
         * @param arg1 the first function argument. Empty values are represented with
         *      @link BlankEval} or {@link MissingArgEval}, never <code>null</code>
         *
         * @param srcRowIndex row index of the cell containing the formula under evaluation
         * @param srcColumnIndex column index of the cell containing the formula under evaluation
         * @return The evaluated result, possibly an {@link ErrorEval}, never <code>null</code>.
         * <b>Note</b> - Excel uses the error code <i>#NUM!</i> instead of IEEE <i>NaN</i>, so when
         * numeric functions evaluate to {@link Double#NaN} be sure to translate the result to {@link
         * ErrorEval#NUM_ERROR}.
         */
        ValueEval EvaluateTwoArrayArgs(ValueEval arg0, ValueEval arg1, int srcRowIndex, int srcColumnIndex,
                                              BiFunction<ValueEval, ValueEval, ValueEval> evalFunc)
        {
            int w1, w2, h1, h2;
            int a1FirstCol = 0, a1FirstRow = 0;
            if (arg0 is AreaEval) {
                AreaEval ae = (AreaEval)arg0;
                w1 = ae.Width;
                h1 = ae.Height;
                a1FirstCol = ae.FirstColumn;
                a1FirstRow = ae.FirstRow;
            } else if (arg0 is RefEval){
                RefEval ref1 = (RefEval)arg0;
                w1 = 1;
                h1 = 1;
                a1FirstCol = ref1.Column;
                a1FirstRow = ref1.Row;
            } else
            {
                w1 = 1;
                h1 = 1;
            }
            int a2FirstCol = 0, a2FirstRow = 0;
            if (arg1 is AreaEval) {
                AreaEval ae = (AreaEval)arg1;
                w2 = ae.Width;
                h2 = ae.Height;
                a2FirstCol = ae.FirstColumn;
                a2FirstRow = ae.FirstRow;
            } else if (arg1 is RefEval){
                RefEval ref1 = (RefEval)arg1;
                w2 = 1;
                h2 = 1;
                a2FirstCol = ref1.Column;
                a2FirstRow = ref1.Row;
            } else
            {
                w2 = 1;
                h2 = 1;
            }

            int width = Math.Max(w1, w2);
            int height = Math.Max(h1, h2);

            ValueEval[] vals = new ValueEval[height * width];

            int idx = 0;
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    ValueEval vA;
                    try
                    {
                        vA = OperandResolver.GetSingleValue(arg0, a1FirstRow + i, a1FirstCol + j);
                    }
                    catch (FormulaParseException e)
                    {
                        vA = ErrorEval.NAME_INVALID;
                    }
                    catch (EvaluationException e)
                    {
                        vA = e.GetErrorEval();
                    }
                    ValueEval vB;
                    try
                    {
                        vB = OperandResolver.GetSingleValue(arg1, a2FirstRow + i, a2FirstCol + j);
                    }
                    catch (FormulaParseException e)
                    {
                        vB = ErrorEval.NAME_INVALID;
                    }
                    catch (EvaluationException e)
                    {
                        vB = e.GetErrorEval();
                    }
                    if (vA is ErrorEval) {
                        vals[idx++] = vA;
                    } else if (vB is ErrorEval) {
                        vals[idx++] = vB;
                    } else
                    {
                        vals[idx++] = evalFunc.apply(vA, vB);
                    }

                }
            }

            if (vals.Length == 1) {
                return vals[0];
            }

            return new CacheAreaEval(srcRowIndex, srcColumnIndex, srcRowIndex + height - 1, srcColumnIndex + width - 1, vals);
        }

        ValueEval EvaluateOneArrayArg(ValueEval[] args, int srcRowIndex, int srcColumnIndex,
                                             java.util.function.Function<ValueEval, ValueEval> evalFunc)
        {
            ValueEval arg0 = args[0];

            int w1, w2, h1, h2;
            int a1FirstCol = 0, a1FirstRow = 0;
            if (arg0 is AreaEval) {
                AreaEval ae = (AreaEval)arg0;
                w1 = ae.Width;
                h1 = ae.Height;
                a1FirstCol = ae.FirstColumn;
                a1FirstRow = ae.FirstRow;
            } else if (arg0 is RefEval) {
                RefEval ref1 = (RefEval)arg0;
                w1 = 1;
                h1 = 1;
                a1FirstCol = ref1.Column;
                a1FirstRow = ref1.Row;
            } else
            {
                w1 = 1;
                h1 = 1;
            }
            w2 = 1;
            h2 = 1;

            int width = Math.Max(w1, w2);
            int height = Math.Max(h1, h2);

            ValueEval[] vals = new ValueEval[height * width];

            int idx = 0;
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    ValueEval vA;
                    try
                    {
                        vA = OperandResolver.GetSingleValue(arg0, a1FirstRow + i, a1FirstCol + j);
                    }
                    catch (FormulaParseException e)
                    {
                        vA = ErrorEval.NAME_INVALID;
                    }
                    catch (EvaluationException e)
                    {
                        vA = e.GetErrorEval();
                    }
                    vals[idx++] = evalFunc.apply(vA);
                }
            }

            if (vals.Length == 1)
            {
                return vals[0];
            }

            return new CacheAreaEval(srcRowIndex, srcColumnIndex, srcRowIndex + height - 1, srcColumnIndex + width - 1, vals);

        }
    }
}
