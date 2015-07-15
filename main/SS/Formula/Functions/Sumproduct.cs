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


using NPOI.Util;

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula;


    /*
     * Implementation for the Excel function SUMPRODUCT<p/>
     * 
     * Syntax : <br/>
     *  SUMPRODUCT ( array1[, array2[, array3[, ...]]])
     *    <table border="0" cellpAdding="1" cellspacing="0" summary="Parameter descriptions">
     *      <tr><th>array1, ... arrayN </th><td>typically area references, 
     *      possibly cell references or scalar values</td></tr>
     *    </table><br/>
     *    
     * Let A<b>n</b><sub>(<b>i</b>,<b>j</b>)</sub> represent the element in the <b>i</b>th row <b>j</b>th column 
     * of the <b>n</b>th array<br/>   
     * Assuming each array has the same dimensions (W, H), the result Is defined as:<br/>    
     * SUMPRODUCT = &Sigma;<sub><b>i</b>: 1..H</sub>  
     * 	(  &Sigma;<sub><b>j</b>: 1..W</sub>  
     * 	  (  &Pi;<sub><b>n</b>: 1..N</sub> 
     * 			A<b>n</b><sub>(<b>i</b>,<b>j</b>)</sub> 
     *    ) 
     *  ) 
     * 
     * @author Josh Micich
     */
    public class Sumproduct : Function
    {

        public ValueEval Evaluate(ValueEval[] args, int srcCellRow, int srcCellCol)
        {

            int maxN = args.Length;

            if (maxN < 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            ValueEval firstArg = args[0];
            try
            {
                if (firstArg is NumericValueEval)
                {
                    return EvaluateSingleProduct(args);
                }
                if (firstArg is RefEval)
                {
                    return EvaluateSingleProduct(args);
                }
                if (firstArg is TwoDEval)
                {
                    TwoDEval ae = (TwoDEval)firstArg;
                    if (ae.IsRow && ae.IsColumn)
                    {
                        return EvaluateSingleProduct(args);
                    }
                    return EvaluateAreaSumProduct(args);
                }
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            throw new RuntimeException("Invalid arg type for SUMPRODUCT: ("
                    + firstArg.GetType().Name + ")");
        }

        private ValueEval EvaluateSingleProduct(ValueEval[] evalArgs)
        {
            int maxN = evalArgs.Length;

            double term = 1D;
            for (int n = 0; n < maxN; n++)
            {
                double val = GetScalarValue(evalArgs[n]);
                term *= val;
            }
            return new NumberEval(term);
        }
        private static double GetScalarValue(ValueEval arg)
        {

            ValueEval eval;
            if (arg is RefEval)
            {
                RefEval re = (RefEval)arg;
                if (re.NumberOfSheets > 1)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                eval = re.GetInnerValueEval(re.FirstSheetIndex);
            }
            else
            {
                eval = arg;
            }

            if (eval == null)
            {
                throw new ArgumentException("parameter may not be null");
            }
            if (eval is AreaEval)
            {
                AreaEval ae = (AreaEval)eval;
                // an area ref can work as a scalar value if it is 1x1
                if (!ae.IsColumn || !ae.IsRow)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                eval = ae.GetRelativeValue(0, 0);
            }

            if (!(eval is ValueEval))
            {
                throw new ArgumentException("Unexpected value eval class ("
                        + eval.GetType().Name + ")");
            }

            return GetProductTerm((ValueEval)eval, true);
        }
        private ValueEval EvaluateAreaSumProduct(ValueEval[] evalArgs)
        {
            int maxN = evalArgs.Length;
            AreaEval[] args = new AreaEval[maxN];
            try
            {
                Array.Copy(evalArgs, 0, args, 0, maxN);
            }
            catch (Exception)
            {
                // one of the other args was not an AreaRef
                return ErrorEval.VALUE_INVALID;
            }


            AreaEval firstArg = args[0];

            int height = firstArg.LastRow - firstArg.FirstRow + 1;
            int width = firstArg.LastColumn - firstArg.FirstColumn + 1; // TODO - junit

            // first check dimensions
            if (!AreasAllSameSize(args, height, width))
            {
                // normally this results in #VALUE!, 
                // but errors in individual cells take precedence
                for (int i = 1; i < args.Length; i++)
                {
                    ThrowFirstError(args[i]);
                }
                return ErrorEval.VALUE_INVALID;
            }
            double acc = 0;

            for (int rrIx = 0; rrIx < height; rrIx++)
            {
                for (int rcIx = 0; rcIx < width; rcIx++)
                {
                    double term = 1D;
                    for (int n = 0; n < maxN; n++)
                    {
                        double val = GetProductTerm(args[n].GetRelativeValue(rrIx, rcIx), false);
                        term *= val;
                    }
                    acc += term;
                }
            }

            return new NumberEval(acc);
        }

        private static void ThrowFirstError(TwoDEval areaEval)
        {
            int height = areaEval.Height;
            int width = areaEval.Width;
            for (int rrIx = 0; rrIx < height; rrIx++)
            {
                for (int rcIx = 0; rcIx < width; rcIx++)
                {
                    ValueEval ve = areaEval.GetValue(rrIx, rcIx);
                    if (ve is ErrorEval)
                    {
                        throw new EvaluationException((ErrorEval)ve);
                    }
                }
            }
        }
        private static bool AreasAllSameSize(TwoDEval[] args, int height, int width)
        {
            for (int i = 0; i < args.Length; i++)
            {
                TwoDEval areaEval = args[i];
                // check that height and width match
                if (areaEval.Height != height)
                {
                    return false;
                }
                if (areaEval.Width != width)
                {
                    return false;
                }
            }
            return true;
        }
        /**
         * Determines a <c>double</c> value for the specified <c>ValueEval</c>. 
         * @param IsScalarProduct <c>false</c> for SUMPRODUCTs over area refs.
         * @throws EvalEx if <c>ve</c> represents an error value.
         * <p/>
         * Note - string values and empty cells are interpreted differently depending on 
         * <c>isScalarProduct</c>.  For scalar products, if any term Is blank or a string, the
         * error (#VALUE!) Is raised.  For area (sum)products, if any term Is blank or a string, the
         * result Is zero.
         */
        private static double GetProductTerm(ValueEval ve, bool IsScalarProduct)
        {

            if (ve is BlankEval || ve == null)
            {
                // TODO - shouldn't BlankEval.INSTANCE be used always instead of null?
                // null seems to occur when the blank cell Is part of an area ref (but not reliably)
                if (IsScalarProduct)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                return 0;
            }

            if (ve is ErrorEval)
            {
                throw new EvaluationException((ErrorEval)ve);
            }
            if (ve is StringEval)
            {
                if (IsScalarProduct)
                {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
                // Note for area SUMPRODUCTs, string values are interpreted as zero
                // even if they would Parse as valid numeric values
                return 0;
            }
            if (ve is NumericValueEval)
            {
                NumericValueEval nve = (NumericValueEval)ve;
                return nve.NumberValue;
            }
            throw new RuntimeException("Unexpected value eval class ("
                    + ve.GetType().Name + ")");
        }
    }
}