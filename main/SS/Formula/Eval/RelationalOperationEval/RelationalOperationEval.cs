/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 10, 2005
 *
 */
namespace NPOI.SS.Formula.Eval
{
    using System;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Util;

    //public class RelationalValues
    //{
    //    public double[] ds = new Double[2];
    //    public bool[] bs = new bool[2];
    //    public string[] ss = new String[3];
    //    public ErrorEval ee = null;
    //}
    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo Dot com &gt;
     *
     */
    public abstract class RelationalOperationEval : Fixed2ArgFunction, ArrayFunction
    {
        private static int DoCompare(ValueEval va, ValueEval vb)
        {
            // special cases when one operand is blank or missing
            if (va == BlankEval.instance|| va is MissingArgEval)
            {
                return CompareBlank(vb);
            }
            if (vb == BlankEval.instance || vb is MissingArgEval)
            {
                return -CompareBlank(va);
            }

            if (va is BoolEval bA)
            {
                if (vb is BoolEval bB)
                {
                    if (bA.BooleanValue == bB.BooleanValue)
                    {
                        return 0;
                    }
                    return bA.BooleanValue ? 1 : -1;
                }
                return 1;
            }
            if (vb is BoolEval)
            {
                return -1;
            }
            if (va is StringEval sA)
            {
                if (vb is StringEval sB)
                {
                    return string.Compare(sA.StringValue, sB.StringValue, StringComparison.OrdinalIgnoreCase);
                }
                return 1;
            }
            if (vb is StringEval)
            {
                return -1;
            }
            if (va is NumberEval nA)
            {
                if (vb is NumberEval nB)
                {
                    if (nA.NumberValue == nB.NumberValue)
                    {
                        // Excel considers -0.0 == 0.0 which is different to Double.compare()
                        return 0;
                    }
                    return NumberComparer.Compare(nA.NumberValue, nB.NumberValue);
                }
            }
            throw new ArgumentException("Bad operand types (" + va.GetType().Name + "), ("
                    + vb.GetType().Name + ")");
        }
        private static int CompareBlank(ValueEval v)
        {
            if (v == BlankEval.instance|| v is MissingArgEval)
            {
                return 0;
            }
            if (v is BoolEval boolEval)
            {
                return boolEval.BooleanValue ? -1 : 0;
            }
            if (v is NumberEval ne)
            {
                //return ne.NumberValue.CompareTo(0.0);
                return NumberComparer.Compare(0.0, ne.NumberValue);
            }
            if (v is StringEval se)
            {
                return se.StringValue.Length < 1 ? 0 : -1;
            }
            throw new ArgumentException("bad value class (" + v.GetType().Name + ")");
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {

            ValueEval vA;
            ValueEval vB;
            try
            {
                vA = OperandResolver.GetSingleValue(arg0, srcRowIndex, srcColumnIndex);
                vB = OperandResolver.GetSingleValue(arg1, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            int cmpResult = DoCompare(vA, vB);
            bool result = ConvertComparisonResult(cmpResult);
            return BoolEval.ValueOf(result);
        }
        public ValueEval EvaluateArray(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            ValueEval arg0 = args[0];
            ValueEval arg1 = args[1];

            int w1, w2, h1, h2;
            int a1FirstCol = 0, a1FirstRow = 0;
            if (arg0 is AreaEval eval)
            {
                w1 = eval.Width;
                h1 = eval.Height;
                a1FirstCol = eval.FirstColumn;
                a1FirstRow = eval.FirstRow;
            }
            else if (arg0 is RefEval ref1)
            {
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
            if (arg1 is AreaEval ae)
            {
                w2 = ae.Width;
                h2 = ae.Height;
                a2FirstCol = ae.FirstColumn;
                a2FirstRow = ae.FirstRow;
            }
            else if (arg1 is RefEval ref1)
            {
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
            for (int i = 0; i < height; i++)
            {
                for (int j = 0; j < width; j++)
                {
                    ValueEval vA;
                    try
                    {
                        vA = OperandResolver.GetSingleValue(arg0, a1FirstRow + i, a1FirstCol + j);
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
                    catch (EvaluationException e)
                    {
                        vB = e.GetErrorEval();
                    }
                    if (vA is ErrorEval)
                    {
                        vals[idx++] = vA;
                    }
                    else if (vB is ErrorEval)
                    {
                        vals[idx++] = vB;
                    }
                    else
                    {
                        int cmpResult = DoCompare(vA, vB);
                        bool result = ConvertComparisonResult(cmpResult);
                        vals[idx++] = BoolEval.ValueOf(result);
                    }

                }
            }

            if (vals.Length == 1)
            {
                return vals[0];
            }

            return new CacheAreaEval(srcRowIndex, srcColumnIndex, srcRowIndex + height - 1, srcColumnIndex + width - 1, vals);
        }
        public abstract bool ConvertComparisonResult(int cmpResult);


        public static Function EqualEval = new EqualEval();
        public static Function NotEqualEval = new NotEqualEval();
        public static Function LessEqualEval = new LessEqualEval();
        public static Function LessThanEval = new LessThanEval();
        public static Function GreaterEqualEval = new GreaterEqualEval();
        public static Function GreaterThanEval = new GreaterThanEval();
    }
}