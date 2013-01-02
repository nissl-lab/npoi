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
    using System.Text;

    /**
     * Implementation for Excel function INDIRECT<p/>
     * 
     * INDIRECT() returns the cell or area reference denoted by the text argument.<p/> 
     * 
     * <b>Syntax</b>:<br/>
     * <b>INDIRECT</b>(<b>ref_text</b>,isA1Style)<p/>
     * 
     * <b>ref_text</b> a string representation of the desired reference as it would normally be written
     * in a cell formula.<br/>
     * <b>isA1Style</b> (default TRUE) specifies whether the ref_text should be interpreted as A1-style
     * or R1C1-style.
     * 
     * 
     * @author Josh Micich
     */
    public class Indirect : FreeRefFunction
    {

        public static FreeRefFunction instance = new Indirect();

        private Indirect()
        {
            // enforce singleton
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            if (args.Length < 1)
            {
                return ErrorEval.VALUE_INVALID;
            }

            bool isA1style;
            String text;
            try
            {
                ValueEval ve = OperandResolver.GetSingleValue(args[0], ec.RowIndex, ec
                        .ColumnIndex);
                text = OperandResolver.CoerceValueToString(ve);
                switch (args.Length)
                {
                    case 1:
                        isA1style = true;
                        break;
                    case 2:
                        isA1style = EvaluateBooleanArg(args[1], ec);
                        break;
                    default:
                        return ErrorEval.VALUE_INVALID;
                }
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }

            return EvaluateIndirect(ec, text, isA1style);
        }

        private static bool EvaluateBooleanArg(ValueEval arg, OperationEvaluationContext ec)
        {
            ValueEval ve = OperandResolver.GetSingleValue(arg, ec.RowIndex, ec.ColumnIndex);

            if (ve == BlankEval.instance || ve == MissingArgEval.instance)
            {
                return false;
            }
            // numeric quantities follow standard bool conversion rules
            // for strings, only "TRUE" and "FALSE" (case insensitive) are valid
            return (bool)OperandResolver.CoerceValueToBoolean(ve, false);
        }

        private static ValueEval EvaluateIndirect(OperationEvaluationContext ec, String text,
                bool isA1style)
        {
            // Search backwards for '!' because sheet names can contain '!'
            int plingPos = text.LastIndexOf('!');

            String workbookName;
            String sheetName;
            String refText; // whitespace around this Gets Trimmed OK
            if (plingPos < 0)
            {
                workbookName = null;
                sheetName = null;
                refText = text;
            }
            else
            {
                String[] parts = ParseWorkbookAndSheetName(text.Substring(0, plingPos));
                if (parts == null)
                {
                    return ErrorEval.REF_INVALID;
                }
                workbookName = parts[0];
                sheetName = parts[1];
                refText = text.Substring(plingPos + 1);
            }

            String refStrPart1;
            String refStrPart2;

            int colonPos = refText.IndexOf(':');
            if (colonPos < 0)
            {
                refStrPart1 = refText.Trim();
                refStrPart2 = null;
            }
            else
            {
                refStrPart1 = refText.Substring(0, colonPos).Trim();
                refStrPart2 = refText.Substring(colonPos + 1).Trim();
            }
            return ec.GetDynamicReference(workbookName, sheetName, refStrPart1, refStrPart2, isA1style);
        }

        /**
         * @return array of length 2: {workbookName, sheetName,}.  Second element will always be
         * present.  First element may be null if sheetName is unqualified.
         * Returns <code>null</code> if text cannot be parsed.
         */
        private static String[] ParseWorkbookAndSheetName(string text)
        {
            int lastIx = text.Length - 1;
            if (lastIx < 0)
            {
                return null;
            }
            if (CanTrim(text))
            {
                return null;
            }
            char firstChar = text[0];
            if (Char.IsWhiteSpace(firstChar))
            {
                return null;
            }
            if (firstChar == '\'')
            {
                // workbookName or sheetName needs quoting
                // quotes go around both
                if (text[lastIx] != '\'')
                {
                    return null;
                }
                firstChar = text[1];
                if (Char.IsWhiteSpace(firstChar))
                {
                    return null;
                }
                String wbName;
                int sheetStartPos;
                if (firstChar == '[')
                {
                    int rbPos = text.ToString().LastIndexOf(']');
                    if (rbPos < 0)
                    {
                        return null;
                    }
                    wbName = UnescapeString(text.Substring(2, rbPos - 2));
                    if (wbName == null || CanTrim(wbName))
                    {
                        return null;
                    }
                    sheetStartPos = rbPos + 1;
                }
                else
                {
                    wbName = null;
                    sheetStartPos = 1;
                }

                // else - just sheet name
                String sheetName = UnescapeString(text.Substring(sheetStartPos, lastIx - sheetStartPos));
                if (sheetName == null)
                { // note - when quoted, sheetName can
                    // start/end with whitespace
                    return null;
                }
                return new String[] { wbName, sheetName, };
            }

            if (firstChar == '[')
            {
                int rbPos = text.ToString().LastIndexOf(']');
                if (rbPos < 0)
                {
                    return null;
                }
                string wbName = text.Substring(1, rbPos - 1);
                if (CanTrim(wbName))
                {
                    return null;
                }
                string sheetName = text.Substring(rbPos + 1);
                if (CanTrim(sheetName))
                {
                    return null;
                }
                return new String[] { wbName.ToString(), sheetName.ToString(), };
            }
            // else - just sheet name
            return new String[] { null, text.ToString(), };
        }

        /**
         * @return <code>null</code> if there is a syntax error in any escape sequence
         * (the typical syntax error is a single quote character not followed by another).
         */
        private static String UnescapeString(string text)
        {
            int len = text.Length;
            StringBuilder sb = new StringBuilder(len);
            int i = 0;
            while (i < len)
            {
                char ch = text[i];
                if (ch == '\'')
                {
                    // every quote must be followed by another
                    i++;
                    if (i >= len)
                    {
                        return null;
                    }
                    ch = text[i];
                    if (ch != '\'')
                    {
                        return null;
                    }
                }
                sb.Append(ch);
                i++;
            }
            return sb.ToString();
        }

        private static bool CanTrim(string text)
        {
            int lastIx = text.Length - 1;
            if (lastIx < 0)
            {
                return false;
            }
            if (Char.IsWhiteSpace(text[0]))
            {
                return true;
            }
            if (Char.IsWhiteSpace(text[lastIx]))
            {
                return true;
            }
            return false;
        }
    }
}