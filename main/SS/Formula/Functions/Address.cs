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

namespace NPOI.SS.Formula.Functions
{
    using NPOI.SS.Formula.Eval;
    using System;
    using NPOI.SS.Util;
    using System.Text;
    /**
     * Creates a text reference as text, given specified row and column numbers.
     *
     * @author Aniket Banerjee (banerjee@google.com)
     */
    public class Address : Function
    {
        public const int REF_ABSOLUTE = 1;
        public const int REF_ROW_ABSOLUTE_COLUMN_RELATIVE = 2;
        public const int REF_ROW_RELATIVE_RELATIVE_ABSOLUTE = 3;
        public const int REF_RELATIVE = 4;

        public ValueEval Evaluate(ValueEval[] args, int srcRowIndex,
                                  int srcColumnIndex)
        {
            if (args.Length < 2 || args.Length > 5)
            {
                return ErrorEval.VALUE_INVALID;
            }
            try
            {
                bool pAbsRow, pAbsCol;

                int row = (int)NumericFunction.SingleOperandEvaluate(args[0], srcRowIndex, srcColumnIndex);
                int col = (int)NumericFunction.SingleOperandEvaluate(args[1], srcRowIndex, srcColumnIndex);

                int refType;
                if (args.Length > 2 && args[2] != MissingArgEval.instance)
                {
                    refType = (int)NumericFunction.SingleOperandEvaluate(args[2], srcRowIndex, srcColumnIndex);
                }
                else
                {
                    refType = REF_ABSOLUTE; // this is also the default if parameter is not given
                }
                switch (refType)
                {
                    case REF_ABSOLUTE:
                        pAbsRow = true;
                        pAbsCol = true;
                        break;
                    case REF_ROW_ABSOLUTE_COLUMN_RELATIVE:
                        pAbsRow = true;
                        pAbsCol = false;
                        break;
                    case REF_ROW_RELATIVE_RELATIVE_ABSOLUTE:
                        pAbsRow = false;
                        pAbsCol = true;
                        break;
                    case REF_RELATIVE:
                        pAbsRow = false;
                        pAbsCol = false;
                        break;
                    default:
                        throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }

                bool a1;
                if (args.Length > 3)
                {
                    ValueEval ve = OperandResolver.GetSingleValue(args[3], srcRowIndex, srcColumnIndex);
                    // TODO R1C1 style is not yet supported
                    a1 = ve == MissingArgEval.instance ? true : OperandResolver.CoerceValueToBoolean(ve, false).Value;
                }
                else
                {
                    a1 = true;
                }

                String sheetName;
                if (args.Length == 5)
                {
                    ValueEval ve = OperandResolver.GetSingleValue(args[4], srcRowIndex, srcColumnIndex);
                    sheetName = ve == MissingArgEval.instance ? null : OperandResolver.CoerceValueToString(ve);
                }
                else
                {
                    sheetName = null;
                }

                CellReference ref1 = new CellReference(row - 1, col - 1, pAbsRow, pAbsCol);
                StringBuilder sb = new StringBuilder(32);
                if (sheetName != null)
                {
                    SheetNameFormatter.AppendFormat(sb, sheetName);
                    sb.Append('!');
                }
                sb.Append(ref1.FormatAsString());

                return new StringEval(sb.ToString());

            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
        }
    }
}