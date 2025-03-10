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

using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.Formula.Eval;
using NPOI.Util;

namespace NPOI.SS.Formula.Functions
{
    /**
     * Implementation for the Excel function SUBTOTAL<p>
     *
     * <b>Syntax :</b> <br/>
     *  SUBTOTAL ( <b>functionCode</b>, <b>ref1</b>, ref2 ... ) <br/>
     *    <table border="1" cellpadding="1" cellspacing="0" summary="Parameter descriptions">
     *      <tr><td><b>functionCode</b></td><td>(1-11) Selects the underlying aggregate function to be used (see table below)</td></tr>
     *      <tr><td><b>ref1</b>, ref2 ...</td><td>Arguments to be passed to the underlying aggregate function</td></tr>
     *    </table><br/>
     * </p>
     *
     *  <table border="1" cellpadding="1" cellspacing="0" summary="Parameter descriptions">
     *      <tr><th>functionCode</th><th>Aggregate Function</th></tr>
     *      <tr align='center'><td>1</td><td>AVERAGE</td></tr>
     *      <tr align='center'><td>2</td><td>COUNT</td></tr>
     *      <tr align='center'><td>3</td><td>COUNTA</td></tr>
     *      <tr align='center'><td>4</td><td>MAX</td></tr>
     *      <tr align='center'><td>5</td><td>MIN</td></tr>
     *      <tr align='center'><td>6</td><td>PRODUCT</td></tr>
     *      <tr align='center'><td>7</td><td>STDEV</td></tr>
     *      <tr align='center'><td>8</td><td>STDEVP *</td></tr>
     *      <tr align='center'><td>9</td><td>SUM</td></tr>
     *      <tr align='center'><td>10</td><td>VAR *</td></tr>
     *      <tr align='center'><td>11</td><td>VARP *</td></tr>
     *      <tr align='center'><td>101-111</td><td>*</td></tr>
     *  </table><br/>
     * * Not implemented in POI yet. Functions 101-111 are the same as functions 1-11 but with
     * the option 'ignore hidden values'.
     * <p/>
     *
     * @author Paul Tomlin &lt; pault at bulk sms dot com &gt;
     */
    public class Subtotal : Function
    {

        private static Function FindFunction(int functionCode)
        {
            //Function func;
            switch (functionCode)
            {
                case 1: return AggregateFunction.SubtotalInstance(AggregateFunction.AVERAGE, true);
                case 2: return Count.SubtotalInstance(true);
                case 3: return Counta.SubtotalInstance(true);
                case 4: return AggregateFunction.SubtotalInstance(AggregateFunction.MAX, true);
                case 5: return AggregateFunction.SubtotalInstance(AggregateFunction.MIN, true);
                case 6: return AggregateFunction.SubtotalInstance(AggregateFunction.PRODUCT, true);
                case 7: return AggregateFunction.SubtotalInstance(AggregateFunction.STDEV, true);
                case 8: return AggregateFunction.SubtotalInstance(AggregateFunction.STDEVP, true);
                case 9: return AggregateFunction.SubtotalInstance(AggregateFunction.SUM, true);
                case 10: return AggregateFunction.SubtotalInstance(AggregateFunction.VAR, true);
                case 11: return AggregateFunction.SubtotalInstance(AggregateFunction.VARP, true);
                case 101: return AggregateFunction.SubtotalInstance(AggregateFunction.AVERAGE, false);
                case 102: return Count.SubtotalInstance(false);
                case 103: return Counta.SubtotalInstance(false);
                case 104: return AggregateFunction.SubtotalInstance(AggregateFunction.MAX, false);
                case 105: return AggregateFunction.SubtotalInstance(AggregateFunction.MIN, false);
                case 106: return AggregateFunction.SubtotalInstance(AggregateFunction.PRODUCT, false);
                case 107: return AggregateFunction.SubtotalInstance(AggregateFunction.STDEV, false);
                case 108: return AggregateFunction.SubtotalInstance(AggregateFunction.STDEVP, false);
                case 109: return AggregateFunction.SubtotalInstance(AggregateFunction.SUM, false);
                case 110: return AggregateFunction.SubtotalInstance(AggregateFunction.VAR, false);
                case 111: return AggregateFunction.SubtotalInstance(AggregateFunction.VARP, false);
            }
            
            throw EvaluationException.InvalidValue();
        }

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            int nInnerArgs = args.Length - 1; // -1: first arg is used to select from a basic aggregate function
            if (nInnerArgs < 1)
            {
                return ErrorEval.VALUE_INVALID;
            }

            Function innerFunc;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(args[0], srcRowIndex, srcColumnIndex);
                int functionCode = OperandResolver.CoerceValueToInt(ve);
                innerFunc = FindFunction(functionCode);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            // ignore the first arg, this is the function-type, we check for the length above
            IList<ValueEval> list = new List<ValueEval>(Arrays.AsList(args).GetRange(1, args.Length - 1));
            IEnumerator<ValueEval> it = list.GetEnumerator();
            // See https://support.office.com/en-us/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939
            // "If there are other subtotals within ref1, ref2,... (or nested subtotals), these nested subtotals are ignored to avoid double counting."
            // For array references it is handled in1 other evaluation steps, but we need to handle this here for references to subtotal-functions
            IList<ValueEval> toRemove = new List<ValueEval>();
            while (it.MoveNext())
            {
                ValueEval eval = it.Current;
                if (eval is LazyRefEval lazyRefEval)
                {
                    if (lazyRefEval.IsSubTotal)
                    {
                        toRemove.Add(lazyRefEval);
                    }
                }
            }

            foreach (var x in toRemove)
                list.Remove(x);

            return innerFunc.Evaluate(list.ToArray(), srcRowIndex, srcColumnIndex);

        }
    }
}