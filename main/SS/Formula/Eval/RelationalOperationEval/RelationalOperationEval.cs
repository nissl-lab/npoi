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
    public abstract class RelationalOperationEval : Fixed2ArgFunction
    {
        private static int DoCompare(ValueEval va, ValueEval vb)
        {
            // special cases when one operand is blank
            if (va == BlankEval.instance)
            {
                return CompareBlank(vb);
            }
            if (vb == BlankEval.instance)
            {
                return -CompareBlank(va);
            }

            if (va is BoolEval)
            {
                if (vb is BoolEval)
                {
                    BoolEval bA = (BoolEval)va;
                    BoolEval bB = (BoolEval)vb;
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
            if (va is StringEval)
            {
                if (vb is StringEval)
                {
                    StringEval sA = (StringEval)va;
                    StringEval sB = (StringEval)vb;
                    return string.Compare(sA.StringValue, sB.StringValue, StringComparison.OrdinalIgnoreCase);
                }
                return 1;
            }
            if (vb is StringEval)
            {
                return -1;
            }
            if (va is NumberEval)
            {
                if (vb is NumberEval)
                {
                    NumberEval nA = (NumberEval)va;
                    NumberEval nB = (NumberEval)vb;
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
        private static int CompareBlank(ValueEval v) {
		    if (v == BlankEval.instance) {
			    return 0;
		    }
		    if (v is BoolEval) {
			    BoolEval boolEval = (BoolEval) v;
			    return boolEval.BooleanValue ? -1 : 0;
		    }
		    if (v is NumberEval) {
			    NumberEval ne = (NumberEval) v;
                //return ne.NumberValue.CompareTo(0.0);
                return NumberComparer.Compare(0.0, ne.NumberValue);
		    }
		    if (v is StringEval) {
			    StringEval se = (StringEval) v;
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

        public abstract bool ConvertComparisonResult(int cmpResult);


        public static Function EqualEval = new EqualEval();
        public static Function NotEqualEval = new NotEqualEval();
        public static Function LessEqualEval = new LessEqualEval();
        public static Function LessThanEval = new LessThanEval();
        public static Function GreaterEqualEval = new GreaterEqualEval();
        public static Function GreaterThanEval = new GreaterThanEval();
    }
}