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
 * Created on May 15, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula;
    using System.Globalization;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * Here are the general rules concerning bool functions:
     * 
     * - Blanks are not either true or false
     * - Strings are not either true or false (even strings "true" or "TRUE" or "0" etc.)
     * - Numbers: 0 Is false. Any other number Is TRUE.
     * - References are Evaluated and above rules apply.
     * - Areas: Individual cells in area are Evaluated and Checked to 
     * see if they are blanks, strings etc.
     */
    public abstract class BooleanFunction : Function
    {
        protected abstract bool InitialResultValue { get; }
        protected abstract bool PartialEvaluate(bool cumulativeResult, bool currentValue);


        private bool Calculate(ValueEval[] args)
        {

            bool result = InitialResultValue;
            bool atleastOneNonBlank = false;
            bool? tempVe;
            /*
             * Note: no short-circuit bool loop exit because any ErrorEvals will override the result
             */
            for (int i = 0, iSize = args.Length; i < iSize; i++)
            {
                ValueEval arg = args[i];
                if (arg is TwoDEval)
                {
                    TwoDEval ae = (TwoDEval)arg;
                    int height = ae.Height;
                    int width = ae.Width;
                    for (int rrIx = 0; rrIx < height; rrIx++)
                    {
                        for (int rcIx = 0; rcIx < width; rcIx++)
                        {
                            ValueEval ve = ae.GetValue(rrIx, rcIx);
                            tempVe = OperandResolver.CoerceValueToBoolean(ve, true);
                            if (tempVe != null)
                            {
                                result = PartialEvaluate(result, Convert.ToBoolean(tempVe, CultureInfo.InvariantCulture));
                                atleastOneNonBlank = true;
                            }
                        }
                    }
                    continue;
                }

                if (arg is RefEval)
                {
                    ValueEval ve = ((RefEval)arg).InnerValueEval;
                    tempVe = OperandResolver.CoerceValueToBoolean(ve, true);
                }
                else if (arg is ValueEval)
                {
                    ValueEval ve = (ValueEval)arg;
                    tempVe = OperandResolver.CoerceValueToBoolean(ve, false);
                }
                else if (arg == MissingArgEval.instance)
                {
                    tempVe = null; // you can leave out parameters, they are simply ignored
                }
                else
                {
                    throw new InvalidOperationException("Unexpected eval (" + arg.GetType().Name + ")");
                }


                if (tempVe != null)
                {
                    result = PartialEvaluate(result, Convert.ToBoolean(tempVe, CultureInfo.InvariantCulture));
                    atleastOneNonBlank = true;
                }
            }

            if (!atleastOneNonBlank)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            return result;
        }

        public ValueEval Evaluate(ValueEval[] args, int srcRow, int srcCol)
        {
            if (args.Length < 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            bool boolResult;
            try
            {
                boolResult = Calculate(args);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return BoolEval.ValueOf(boolResult);
        }
    }
}