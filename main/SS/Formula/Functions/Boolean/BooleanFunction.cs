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

namespace NPOI.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula;
    using System.Globalization;

    /**
     * Here are the general rules concerning Boolean functions:
     * <ol>
     * <li> Blanks are ignored (not either true or false) </li>
     * <li> Strings are ignored if part of an area ref or cell ref, otherwise they must be 'true' or 'false'</li>
     * <li> Numbers: 0 is false. Any other number is TRUE </li>
     * <li> Areas: *all* cells in area are evaluated according to the above rules</li>
     * </ol>
     *
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     */
    public abstract class BooleanFunction : Function
    {
        protected abstract bool InitialResultValue { get; }
        protected abstract bool PartialEvaluate(bool cumulativeResult, bool currentValue);


        private bool Calculate(ValueEval[] args)
        {

            bool result = InitialResultValue;
            bool atleastOneNonBlank = false;
           
            /*
             * Note: no short-circuit bool loop exit because any ErrorEvals will override the result
             */
            for (int i = 0, iSize = args.Length; i < iSize; i++)
            {
                bool? tempVe;
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
                    RefEval re = (RefEval)arg;
                    for (int sIx = re.FirstSheetIndex; sIx <= re.LastSheetIndex; sIx++)
                    {
                        ValueEval ve = re.GetInnerValueEval(sIx);
                        tempVe = OperandResolver.CoerceValueToBoolean(ve, true);
                        if (tempVe != null)
                        {
                            result = PartialEvaluate(result, tempVe.Value);
                            atleastOneNonBlank = true;
                        }
                    }
                    continue;
                }
                //else if (arg is ValueEval)
                //{
                //    ValueEval ve = (ValueEval)arg;
                //    tempVe = OperandResolver.CoerceValueToBoolean(ve, false);
                //}
                if (arg == MissingArgEval.instance)
                {
                    tempVe = null; // you can leave out parameters, they are simply ignored
                }
                else
                {
                    tempVe = OperandResolver.CoerceValueToBoolean(arg, false);
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