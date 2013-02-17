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

namespace NPOI.SS.Formula.Functions
{

    using NPOI.SS.Formula.Eval;

    /**
     * Implementation for the Excel function TIME
     *
     * @author Steven Butler (sebutler @ gmail dot com)
     *
     * Based on POI {@link DateFunc}
     */
    public class TimeFunc : Fixed3ArgFunction
    {

        private const int SECONDS_PER_MINUTE = 60;
        private const int SECONDS_PER_HOUR = 3600;
        private const int HOURS_PER_DAY = 24;
        private const int SECONDS_PER_DAY = HOURS_PER_DAY * SECONDS_PER_HOUR;


        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1,
                ValueEval arg2)
        {
            double result;
            try
            {
                result = Evaluate(EvalArg(arg0, srcRowIndex, srcColumnIndex), EvalArg(arg1, srcRowIndex, srcColumnIndex), EvalArg(arg2, srcRowIndex, srcColumnIndex));
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return new NumberEval(result);
        }
        private int EvalArg(ValueEval arg, int srcRowIndex, int srcColumnIndex)
        {
            if (arg == MissingArgEval.instance)
            {
                return 0;
            }
            ValueEval ev = OperandResolver.GetSingleValue(arg, srcRowIndex, srcColumnIndex);
            // Excel silently tRuncates double values to integers
            return OperandResolver.CoerceValueToInt(ev);
        }
        /**
         * Converts the supplied hours, minutes and seconds to an Excel time value.
         *
         *
         * @param ds array of 3 doubles Containing hours, minutes and seconds.
         * Non-integer inputs are tRuncated to an integer before further calculation
         * of the time value.
         * @return An Excel representation of a time of day.
         * If the time value represents more than a day, the days are Removed from
         * the result, leaving only the time of day component.
         * @throws NPOI.SS.Formula.Eval.EvaluationException
         * If any of the arguments are greater than 32767 or the hours
         * minutes and seconds when combined form a time value less than 0, the function
         * Evaluates to an error.
         */
        private double Evaluate(int hours, int minutes, int seconds)
        {

            if (hours > 32767 || minutes > 32767 || seconds > 32767)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            int totalSeconds = hours * SECONDS_PER_HOUR + minutes * SECONDS_PER_MINUTE + seconds;

            if (totalSeconds < 0)
            {
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            }
            return (totalSeconds % SECONDS_PER_DAY) / (double)SECONDS_PER_DAY;
        }
    }

}