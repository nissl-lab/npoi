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
using NPOI.SS.Formula.Eval;
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
                case 1: return AggregateFunction.SubtotalInstance(AggregateFunction.AVERAGE);
                case 2: return Count.SubtotalInstance();
                case 3: return Counta.SubtotalInstance();
                case 4: return AggregateFunction.SubtotalInstance(AggregateFunction.MAX);
                case 5: return AggregateFunction.SubtotalInstance(AggregateFunction.MIN);
                case 6: return AggregateFunction.SubtotalInstance(AggregateFunction.PRODUCT);
                case 7: return AggregateFunction.SubtotalInstance(AggregateFunction.STDEV);
                case 8: throw new NotImplementedFunctionException("STDEVP");
                case 9: return AggregateFunction.SubtotalInstance(AggregateFunction.SUM);
                case 10: throw new NotImplementedFunctionException("VAR");
                case 11: throw new NotImplementedFunctionException("VARP");
            }
            if (functionCode > 100 && functionCode < 112)
            {
                throw new NotImplementedException("SUBTOTAL - with 'exclude hidden values' option");
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

            ValueEval[] innerArgs = new ValueEval[nInnerArgs];
            Array.Copy(args, 1, innerArgs, 0, nInnerArgs);

            return innerFunc.Evaluate(innerArgs, srcRowIndex, srcColumnIndex);
        }
    }
}