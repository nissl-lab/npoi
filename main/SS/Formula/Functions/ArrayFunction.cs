/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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

using NPOI.SS.Formula;
using NPOI.SS.Formula.Eval;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula.Functions
{
    public interface IArrayFunction
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

    public class ArrayFunction
    {
        /// <summary>
        /// Evaluate an array function with two arguments.
        /// </summary>
        /// <param name="arg0">the first function argument. Empty values are represented with <see cref="BlankEval"/> or <see cref="MissingArgEval"/>, never <c>null</c></param>
        /// <param name="arg1">the second function argument. Empty values are represented with <see cref="BlankEval"/> or <see cref="MissingArgEval"/>, never <c>null</c></param>
        /// <param name="srcRowIndex">row index of the cell containing the formula under evaluation</param>
        /// <param name="srcColumnIndex">column index of the cell containing the formula under evaluation</param>
        /// <param name="evalFunc"></param>
        /// <returns>The evaluated result, possibly an <see cref="ErrorEval"/>, never <c>null</c>.
        /// <b>Note</b> - Excel uses the error code <i>#NUM!</i> instead of IEEE <i>NaN</i>, so when
        /// numeric functions evaluate to <see cref="double.NaN"/> be sure to translate the result to 
        /// <see cref="ErrorEval.NUM_ERROR"/>. </returns>
        public ValueEval EvaluateTwoArrayArgs(ValueEval arg0, ValueEval arg1, int srcRowIndex, int srcColumnIndex,
                                           Func<ValueEval, ValueEval, ValueEval> evalFunc)
        {
            return _evaluateTwoArrayArgs(arg0, arg1, srcRowIndex, srcColumnIndex, evalFunc);
        }

        public ValueEval EvaluateOneArrayArg(ValueEval arg0, int srcRowIndex, int srcColumnIndex,
                                          Func<ValueEval, ValueEval> evalFunc)
        {
            return _evaluateOneArrayArg(arg0, srcRowIndex, srcColumnIndex, evalFunc);
        }

        internal static ValueEval _evaluateTwoArrayArgs(ValueEval arg0, ValueEval arg1, int srcRowIndex, int srcColumnIndex,
            Func<ValueEval, ValueEval, ValueEval> evalFunc)
        {
            int w1, w2, h1, h2;
            int a1FirstCol = 0, a1FirstRow = 0;
            if(arg0 is AreaEval)
            {
                AreaEval ae = (AreaEval)arg0;
                w1 = ae.Width;
                h1 = ae.Height;
                a1FirstCol = ae.FirstColumn;
                a1FirstRow = ae.FirstRow;
            }
            else if(arg0 is RefEval)
            {
                RefEval ref1 = (RefEval)arg0;
                w1 = 1;
                h1 = 1;
                a1FirstCol = ref1.Column;
                a1FirstRow = ref1.Row;
            }
            else
            {
                w1 = 1;
                h1 = 1;
            }
            int a2FirstCol = 0, a2FirstRow = 0;
            if(arg1 is AreaEval)
            {
                AreaEval ae = (AreaEval)arg1;
                w2 = ae.Width;
                h2 = ae.Height;
                a2FirstCol = ae.FirstColumn;
                a2FirstRow = ae.FirstRow;
            }
            else if(arg1 is RefEval)
            {
                RefEval ref1 = (RefEval)arg1;
                w2 = 1;
                h2 = 1;
                a2FirstCol = ref1.Column;
                a2FirstRow = ref1.Row;
            }
            else
            {
                w2 = 1;
                h2 = 1;
            }

            int width = Math.Max(w1, w2);
            int height = Math.Max(h1, h2);

            ValueEval[] vals = new ValueEval[height * width];

            int idx = 0;
            for(int i = 0; i < height; i++)
            {
                for(int j = 0; j < width; j++)
                {
                    ValueEval vA;
                    try
                    {
                        vA = OperandResolver.GetSingleValue(arg0, a1FirstRow + i, a1FirstCol + j);
                    }
                    catch(FormulaParseException e)
                    {
                        vA = ErrorEval.NAME_INVALID;
                    }
                    catch(EvaluationException e)
                    {
                        vA = e.GetErrorEval();
                    }
                    catch(RuntimeException e)
                    {
                        if(e.Message.StartsWith("Don't know how to evaluate name"))
                        {
                            vA = ErrorEval.NAME_INVALID;
                        }
                        else
                        {
                            throw;
                        }
                    }
                    ValueEval vB;
                    try
                    {
                        vB = OperandResolver.GetSingleValue(arg1, a2FirstRow + i, a2FirstCol + j);
                    }
                    catch(FormulaParseException e)
                    {
                        vB = ErrorEval.NAME_INVALID;
                    }
                    catch(EvaluationException e)
                    {
                        vB = e.GetErrorEval();
                    }
                    catch(RuntimeException e)
                    {
                        if(e.Message.StartsWith("Don't know how to evaluate name"))
                        {
                            vB = ErrorEval.NAME_INVALID;
                        }
                        else
                        {
                            throw;
                        }
                    }
                    if(vA is ErrorEval)
                    {
                        vals[idx++] = vA;
                    }
                    else if(vB is ErrorEval)
                    {
                        vals[idx++] = vB;
                    }
                    else
                    {
                        vals[idx++] = evalFunc.Invoke(vA, vB);
                    }

                }
            }

            if(vals.Length == 1)
            {
                return vals[0];
            }

            return new CacheAreaEval(srcRowIndex, srcColumnIndex, srcRowIndex + height - 1, srcColumnIndex + width - 1, vals);
        }


        internal static ValueEval _evaluateOneArrayArg(ValueEval arg0, int srcRowIndex, int srcColumnIndex,
            Func<ValueEval, ValueEval> evalFunc)
        {
            int w1, w2, h1, h2;
            int a1FirstCol = 0, a1FirstRow = 0;
            if(arg0 is AreaEval)
            {
                AreaEval ae = (AreaEval)arg0;
                w1 = ae.Width;
                h1 = ae.Height;
                a1FirstCol = ae.FirstColumn;
                a1FirstRow = ae.FirstRow;
            }
            else if(arg0 is RefEval)
            {
                RefEval ref1 = (RefEval)arg0;
                w1 = 1;
                h1 = 1;
                a1FirstCol = ref1.Column;
                a1FirstRow = ref1.Row;
            }
            else
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
            for(int i = 0; i < height; i++)
            {
                for(int j = 0; j < width; j++)
                {
                    ValueEval vA;
                    try
                    {
                        vA = OperandResolver.GetSingleValue(arg0, a1FirstRow + i, a1FirstCol + j);
                    }
                    catch(FormulaParseException e)
                    {
                        vA = ErrorEval.NAME_INVALID;
                    }
                    catch(EvaluationException e)
                    {
                        vA = e.GetErrorEval();
                    }
                    catch(RuntimeException e)
                    {
                        if(e.Message.StartsWith("Don't know how to evaluate name"))
                        {
                            vA = ErrorEval.NAME_INVALID;
                        }
                        else
                        {
                            throw;
                        }
                    }
                    vals[idx++] = evalFunc.Invoke(vA);
                }
            }

            if(vals.Length == 1)
            {
                return vals[0];
            }

            return new CacheAreaEval(srcRowIndex, srcColumnIndex, srcRowIndex + height - 1, srcColumnIndex + width - 1, vals);
        }
    }
}
