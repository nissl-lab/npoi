/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Formula.PTG
{
    using System;
    
    using NPOI.Util;
    using NPOI.HSSF.UserModel;

    /**
     * @author Daniel Noll (daniel at nuix dot com dot au)
     */
    public class ErrPtg : ScalarConstantPtg
    {

        // convenient access to namespace
        //private static HSSFErrorConstants EC = null;

        /** <b>#NULL!</b>  - Intersection of two cell ranges is empty */
        public static readonly ErrPtg NULL_INTERSECTION = new ErrPtg(HSSFErrorConstants.ERROR_NULL);
        /** <b>#DIV/0!</b> - Division by zero */
        public static readonly ErrPtg DIV_ZERO = new ErrPtg(HSSFErrorConstants.ERROR_DIV_0);
        /** <b>#VALUE!</b> - Wrong type of operand */
        public static readonly ErrPtg VALUE_INVALID = new ErrPtg(HSSFErrorConstants.ERROR_VALUE);
        /** <b>#REF!</b> - Illegal or deleted cell reference */
        public static readonly ErrPtg REF_INVALID = new ErrPtg(HSSFErrorConstants.ERROR_REF);
        /** <b>#NAME?</b> - Wrong function or range name */
        public static readonly ErrPtg NAME_INVALID = new ErrPtg(HSSFErrorConstants.ERROR_NAME);
        /** <b>#NUM!</b> - Value range overflow */
        public static readonly ErrPtg NUM_ERROR = new ErrPtg(HSSFErrorConstants.ERROR_NUM);
        /** <b>#N/A</b> - Argument or function not available */
        public static readonly ErrPtg N_A = new ErrPtg(HSSFErrorConstants.ERROR_NA);

        public static ErrPtg ValueOf(int code)
        {
            switch (code)
            {
                case HSSFErrorConstants.ERROR_DIV_0: return DIV_ZERO;
                case HSSFErrorConstants.ERROR_NA: return N_A;
                case HSSFErrorConstants.ERROR_NAME: return NAME_INVALID;
                case HSSFErrorConstants.ERROR_NULL: return NULL_INTERSECTION;
                case HSSFErrorConstants.ERROR_NUM: return NUM_ERROR;
                case HSSFErrorConstants.ERROR_REF: return REF_INVALID;
                case HSSFErrorConstants.ERROR_VALUE: return VALUE_INVALID;
            }
            throw new InvalidOperationException("Unexpected error code (" + code + ")");
        }

        public const byte sid = 0x1c;
        private const int SIZE = 2;
        private int field_1_error_code;

        /** Creates new ErrPtg */

        public ErrPtg(int errorCode)
        {
            if (!HSSFErrorConstants.IsValidCode(errorCode))
            {
                throw new ArgumentException("Invalid error code (" + errorCode + ")");
            }
            field_1_error_code = errorCode;
        }

        public ErrPtg(ILittleEndianInput in1)
            : this(in1.ReadByte())
        {
            
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
            out1.WriteByte((byte)field_1_error_code);
        }

        public override String ToFormulaString()
        {
            return HSSFErrorConstants.GetText(field_1_error_code);
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public int ErrorCode
        {
            get { return field_1_error_code; }
        }
    }
}