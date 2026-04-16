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
using EFF = Excel.FinancialFunctions;
using NPOI.SS.Formula.Eval;

namespace NPOI.SS.Formula.Functions
{
    internal static class FinancialHelper
    {
        private static readonly DateTime ExcelEpoch = new DateTime(1899, 12, 31);

        internal static EFF.PaymentDue ToPaymentDue(double typeArg)
        {
            return typeArg != 0.0 ? EFF.PaymentDue.BeginningOfPeriod : EFF.PaymentDue.EndOfPeriod;
        }

        internal static EFF.DayCountBasis ToDayCountBasis(int basis)
        {
            switch (basis)
            {
                case 0: return EFF.DayCountBasis.UsPsa30_360;
                case 1: return EFF.DayCountBasis.ActualActual;
                case 2: return EFF.DayCountBasis.Actual360;
                case 3: return EFF.DayCountBasis.Actual365;
                case 4: return EFF.DayCountBasis.Europ30_360;
                default: throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
        }

        internal static EFF.Frequency ToFrequency(int freq)
        {
            switch (freq)
            {
                case 1: return EFF.Frequency.Annual;
                case 2: return EFF.Frequency.SemiAnnual;
                case 4: return EFF.Frequency.Quarterly;
                default: throw new EvaluationException(ErrorEval.NUM_ERROR);
            }
        }

        internal static DateTime FromExcelDate(double serialDate)
        {
            if (serialDate < 1)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            int serial = (int)serialDate;
            // Excel has a Lotus 1-2-3 compatibility bug treating 1900 as a leap year.
            // Serial 60 represents the non-existent Feb 29, 1900. Dates from serial 61 onward
            // are one day ahead of the actual day count from 1899-12-31.
            if (serial >= 61)
                serial -= 1;
            return ExcelEpoch.AddDays(serial);
        }

        internal static double GetDoubleArg(ValueEval[] args, int index, int srcRowIndex, int srcColumnIndex, double defaultValue = double.NaN)
        {
            if (index >= args.Length)
            {
                if (double.IsNaN(defaultValue))
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                return defaultValue;
            }
            ValueEval ve = args[index];
            if (ve == MissingArgEval.instance)
            {
                if (double.IsNaN(defaultValue))
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                return defaultValue;
            }
            ValueEval sv = OperandResolver.GetSingleValue(ve, srcRowIndex, srcColumnIndex);
            return OperandResolver.CoerceValueToDouble(sv);
        }

        internal static int GetIntArg(ValueEval[] args, int index, int srcRowIndex, int srcColumnIndex, int defaultValue = int.MinValue)
        {
            double d = GetDoubleArg(args, index, srcRowIndex, srcColumnIndex,
                defaultValue == int.MinValue ? double.NaN : (double)defaultValue);
            return (int)d;
        }

        internal static DateTime GetDateArg(ValueEval[] args, int index, int srcRowIndex, int srcColumnIndex)
        {
            double serial = GetDoubleArg(args, index, srcRowIndex, srcColumnIndex);
            return FromExcelDate(serial);
        }

        internal static double[] GetArrayArg(ValueEval[] args, int index)
        {
            if (index >= args.Length)
                throw new EvaluationException(ErrorEval.VALUE_INVALID);
            return AggregateFunction.ValueCollector.CollectValues(args[index]);
        }
    }
}
