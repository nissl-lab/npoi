/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */
using System;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;

namespace NPOI.SS.Formula.Functions
{
    /**
 * Implementation for the Excel function WEEKDAY
 *
 * @author Thies Wellpott
 */
    public class WeekdayFunc : Function
    {
        //or:  extends Var1or2ArgFunction {

        public static Function instance = new WeekdayFunc();

        private WeekdayFunc()
        {
            // no fields to initialise
        }

        /* for Var1or2ArgFunction:
        @Override
        public ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0) {
        }

        @Override
        public ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1) {
        }
        */


        /**
         * Perform WEEKDAY(date, returnOption) function.
         * Note: Parameter texts are from German EXCEL-2010 help.
         * Parameters in args[]:
         *  args[0] serialDate
         * EXCEL-date value
         * Standardmaessig ist der 1. Januar 1900 die fortlaufende Zahl 1 und
         * der 1. Januar 2008 die fortlaufende Zahl 39.448, da dieser Tag nach 39.448 Tagen
         * auf den 01.01.1900 folgt.
         * @return Option (optional)
         * Bestimmt den Rueckgabewert:
            1	oder nicht angegeben Zahl 1 (Sonntag) bis 7 (Samstag). Verhaelt sich wie fruehere Microsoft Excel-Versionen.
            2	Zahl 1 (Montag) bis 7 (Sonntag).
            3	Zahl 0 (Montag) bis 6 (Sonntag).
            11	Die Zahlen 1 (Montag) bis 7 (Sonntag)
            12	Die Zahlen 1 (Dienstag) bis 7 (Montag)
            13	Die Zahlen 1 (Mittwoch) bis 7 (Dienstag)
            14	Die Zahlen 1 (Donnerstag) bis 7 (Mittwoch)
            15	Die Zahlen 1 (Freitag) bis 7 (Donnerstag)
            16	Die Zahlen 1 (Samstag) bis 7 (Freitag)
            17	Die Zahlen 1 (Sonntag) bis 7 (Samstag)
         */
        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex)
        {
            try
            {
                if (args.Length < 1 || args.Length > 2)
                {
                    return ErrorEval.VALUE_INVALID;
                }

                // extract first parameter
                ValueEval serialDateVE = OperandResolver.GetSingleValue(args[0], srcRowIndex, srcColumnIndex);
                double serialDate = OperandResolver.CoerceValueToDouble(serialDateVE);
                if (!DateUtil.IsValidExcelDate(serialDate))
                {
                    return ErrorEval.NUM_ERROR;						// EXCEL uses this and no VALUE_ERROR
                }
                DateTime date = DateUtil.GetJavaCalendar(serialDate, false);		// (XXX 1904-windowing not respected)
                int weekday = (int)date.DayOfWeek;		// => sunday = 1, monday = 2, ..., saturday = 7

                // extract second parameter
                int returnOption = 1;					// default value
                if (args.Length == 2)
                {
                    ValueEval ve = OperandResolver.GetSingleValue(args[1], srcRowIndex, srcColumnIndex);
                    if (ve == MissingArgEval.instance || ve == BlankEval.instance)
                    {
                        return ErrorEval.NUM_ERROR;		// EXCEL uses this and no VALUE_ERROR
                    }
                    returnOption = OperandResolver.CoerceValueToInt(ve);
                    if (returnOption == 2)
                    {
                        returnOption = 11;				// both mean the same
                    }
                } // if

                // perform calculation
                double result;
                if (returnOption == 1)
                {
                    result = weekday;
                    // value 2 is handled above (as value 11)
                }
                else if (returnOption == 3)
                {
                    result = (weekday + 6 - 1) % 7;
                }
                else if (returnOption >= 11 && returnOption <= 17)
                {
                    result = (weekday + 6 - (returnOption - 10)) % 7 + 1;		// rotate in the value range 1 to 7
                }
                else
                {
                    return ErrorEval.NUM_ERROR;		// EXCEL uses this and no VALUE_ERROR
                }

                return new NumberEval(result);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        } // evaluate()

    }

}
