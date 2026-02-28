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
    using System;

    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;

    /**
     * <p>Implementation for Excel Oct2Dec() function.</p>
     * <p>
     * Converts an octal number to decimal.
     * </p>
     * <p>
     * <b>Syntax</b>:<br/> <b>Oct2Dec  </b>(<b>number</b> )
     * </p>
     * <p/>
     * Number     is the octal number you want to Convert. Number may not contain more than 10 octal characters (30 bits).
     * The most significant bit of number is the sign bit. The remaining 29 bits are magnitude bits.
     * Negative numbers are represented using two's-complement notation..
     * <p/>
     * If number is not a valid octal number, OCT2DEC returns the #NUM! error value.
     *
     * @author cedric dot walter @ gmail dot com
     */
    public class Oct2Dec : Fixed1ArgFunction, FreeRefFunction
    {

        public static FreeRefFunction instance = new Oct2Dec();

        static int MAX_NUMBER_OF_PLACES = 10;
        static int OCTAL_BASE = 8;

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval numberVE)
        {
            String octal = OperandResolver.CoerceValueToString(numberVE);
            try
            {
                return new NumberEval(BaseNumberUtils.ConvertToDecimal(octal, OCTAL_BASE, MAX_NUMBER_OF_PLACES));
            }
            catch (ArgumentException)
            {
                return ErrorEval.NUM_ERROR;
            }
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length != 1)
            {
                return ErrorEval.VALUE_INVALID;
            }
            return Evaluate(ec.RowIndex, ec.ColumnIndex, args[0]);
        }
    }

}