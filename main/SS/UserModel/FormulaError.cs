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
namespace NPOI.SS.UserModel
{
    using System;

    /**
     * Enumerates error values in SpreadsheetML formula calculations.
     *
     * @author Yegor Kozlov
     */
    public class FormulaError
    {
        /**
         * Intended to indicate when two areas are required to intersect, but do not.
         * <p>Example:
         * In the case of SUM(B1 C1), the space between B1 and C1 is treated as the binary
         * intersection operator, when a comma was intended. end example]
         * </p>
         */
        public static readonly FormulaError NULL = new FormulaError(0x00, "#NULL!");

        /**
         * Intended to indicate when any number, including zero, is divided by zero.
         * Note: However, any error code divided by zero results in that error code.
         */
        public static readonly FormulaError DIV0 = new FormulaError(0x07, "#DIV/0!");

        /**
         * Intended to indicate when an incompatible type argument is passed to a function, or
         * an incompatible type operand is used with an operator.
         * <p>Example:
         * In the case of a function argument, text was expected, but a number was provided
         * </p>
         */
        public static readonly FormulaError VALUE = new FormulaError(0x0F, "#VALUE!");

        /**
         * Intended to indicate when a cell reference is invalid.
         * <p>Example:
         * If a formula Contains a reference to a cell, and then the row or column Containing that cell is deleted,
         * a #REF! error results. If a worksheet does not support 20,001 columns,
         * OFFSET(A1,0,20000) will result in a #REF! error.
         * </p>
         */
        public static readonly FormulaError REF = new FormulaError(0x17, "#REF!");

        /*
         * Intended to indicate when what looks like a name is used, but no such name has been defined.
         * <p>Example:
         * XYZ/3, where XYZ is not a defined name. Total is &amp; A10,
         * where neither Total nor is is a defined name. Presumably, "Total is " & A10
         * was intended. SUM(A1C10), where the range A1:C10 was intended.
         * </p>
         */
        public static readonly FormulaError NAME = new FormulaError(0x1D, "#NAME?");

        /**
         * Intended to indicate when an argument to a function has a compatible type, but has a
         * value that is outside the domain over which that function is defined. (This is known as
         * a domain error.)
         * <p>Example:
         * Certain calls to ASIN, ATANH, FACT, and SQRT might result in domain errors.
         * </p>
         * Intended to indicate that the result of a function cannot be represented in a value of
         * the specified type, typically due to extreme magnitude. (This is known as a range
         * error.)
         * <p>Example: FACT(1000) might result in a range error. </p>
         */
        public static readonly FormulaError NUM = new FormulaError(0x24, "#NUM!");

        /**
         * Intended to indicate when a designated value is not available.
         * <p>Example:
         * Some functions, such as SUMX2MY2, perform a series of operations on corresponding
         * elements in two arrays. If those arrays do not have the same number of elements, then
         * for some elements in the longer array, there are no corresponding elements in the
         * shorter one; that is, one or more values in the shorter array are not available.
         * </p>
         * This error value can be produced by calling the function NA
         */
        public static readonly FormulaError NA = new FormulaError(0x2A, "#N/A");

        private byte type;
        private String repr;

        private FormulaError(int type, String repr)
        {
            this.type = (byte)type;
            this.repr = repr;
            //if(imap==null)
            //    imap = new Dictionary<Byte, FormulaError>();
            //imap.Add(this.Code, this);
            //if (smap == null)
            //    smap = new Dictionary<string, FormulaError>();
            //smap.Add(this.String, this);
        }

        /**
         * @return numeric code of the error
         */
        public byte Code
        {
            get
            {
                return type;
            }
        }

        /**
         * @return string representation of the error
         */
        public String String
        {
            get
            {
                return repr;
            }
        }

        //private static Dictionary<String, FormulaError> smap = null;
        //private static Dictionary<Byte, FormulaError> imap = null;

        public static FormulaError ForInt(byte type)
        {
            //FormulaError err = imap[type];
            //if (err == null) throw new ArgumentException("Unknown error type: " + type);
            //return err;
            switch (type)
            {
                case 0x00: return FormulaError.NULL;
                case 0x07: return FormulaError.DIV0;
                case 0x0F: return FormulaError.VALUE;
                case 0x17: return FormulaError.REF;
                case 0x1D: return FormulaError.NAME;
                case 0x24: return FormulaError.NUM;
                case 0x2A: return FormulaError.NA;
            }
            throw new ArgumentException("Unknown error type: " + type);
        }

        public static FormulaError ForString(String code)
        {
            //FormulaError err = smap[code];
            //if (err == null) throw new ArgumentException("Unknown error code: " + code);
            //return err;
            switch (code)
            {
                case "#NULL!": return FormulaError.NULL;
                case "#DIV/0!": return FormulaError.DIV0;
                case "#VALUE!": return FormulaError.VALUE;
                case "#REF!": return FormulaError.REF;
                case "#NAME?": return FormulaError.NAME;
                case "#NUM!": return FormulaError.NUM;
                case "#N/A": return FormulaError.NA;
            }
            throw new ArgumentException("Unknown error code: " + code);
        }
    }

}

