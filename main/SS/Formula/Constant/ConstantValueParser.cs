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

namespace NPOI.SS.Formula.Constant
{
    using NPOI.Util;
    using System;
    using NPOI.Util.IO;


    /**
     * To support Constant Values (2.5.7) as required by the CRN record.
     * This class is also used for two dimensional arrays which are encoded by 
     * EXTERNALNAME (5.39) records and Array tokens.<p/>
     * 
     * @author Josh Micich
     */
    public class ConstantValueParser
    {
        // note - these (non-combinable) enum values are sparse.
        private const int TYPE_EMPTY = 0;
        private const int TYPE_NUMBER = 1;
        private const int TYPE_STRING = 2;
        private const int TYPE_BOOLEAN = 4;
        private const int TYPE_ERROR_CODE = 16; // TODO - update OOO document to include this value

        private const int TRUE_ENCODING = 1;
        private const int FALSE_ENCODING = 0;

        // TODO - is this the best way to represent 'EMPTY'?
        private static object EMPTY_REPRESENTATION = null;

        private ConstantValueParser()
        {
            // no instances of this class
        }

        public static object[] Parse(LittleEndianInput input, int nValues)
        {
            object[] result = new object[nValues];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = ReadAConstantValue(input);
            }
            return result;
        }

        private static Object ReadAConstantValue(LittleEndianInput input)
        {
            byte grbit = (byte)input.ReadByte();
            switch ((int)grbit)
            {
                case TYPE_EMPTY:
                    input.ReadLong(); // 8 byte 'not used' field
                    return EMPTY_REPRESENTATION;
                case TYPE_NUMBER:
                    return input.ReadDouble();
                case TYPE_STRING:
                    return StringUtil.ReadUnicodeString(input);
                case TYPE_BOOLEAN:
                    return ReadBoolean(input);
                case TYPE_ERROR_CODE:
                    int errCode = input.ReadUShort();
                    // next 6 bytes are unused
                    input.ReadUShort();
                    input.ReadInt();
                    return ErrorConstant.ValueOf(errCode);
            }
            throw new NotSupportedException("Unknown grbit value (" + grbit + ")");
        }

        private static object ReadBoolean(LittleEndianInput input)
        {
            byte val = (byte)input.ReadLong(); // 7 bytes 'not used'
            switch ((int)val)
            {
                case FALSE_ENCODING:
                    return false;
                case TRUE_ENCODING:
                    return true;
            }
            // Don't tolerate unusual boolean encoded values (unless it becomes evident that they occur)
            throw new NotSupportedException("unexpected boolean encoding (" + val + ")");
        }

        public static int GetEncodedSize(object[] values)
        {
            // start with one byte 'type' code for each value
            int result = values.Length * 1;
            for (int i = 0; i < values.Length; i++)
            {
                result += GetEncodedSize(values[i]);
            }
            return result;
        }

        /**
         * @return encoded size without the 'type' code byte
         */
        private static int GetEncodedSize(object obj)
        {
            if (obj == EMPTY_REPRESENTATION)
            {
                return 8;
            }
            Type cls = obj.GetType();

            if (cls == typeof(bool) || cls == typeof(double) || cls == typeof(ErrorConstant))
            {
                return 8;
            }
            String strVal = obj.ToString();
            return StringUtil.GetEncodedSize(strVal);
        }

        public static void Encode(LittleEndianOutput output, object[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                EncodeSingleValue(output, values[i]);
            }
        }

        private static void EncodeSingleValue(LittleEndianOutput output, object value)
        {
            if (value == EMPTY_REPRESENTATION)
            {
                output.WriteByte(TYPE_EMPTY);
                output.WriteLong(0L);
                return;
            }
            if (value is Boolean)
            {
                Boolean bVal = ((Boolean)value);
                output.WriteByte(TYPE_BOOLEAN);
                long longVal = bVal ? 1L : 0L;
                output.WriteLong(longVal);
                return;
            }
            if (value is Double)
            {
                Double dVal = (Double)value;
                output.WriteByte(TYPE_NUMBER);
                output.WriteDouble(dVal);
                return;
            }
            if (value is String)
            {
                String val = (String)value;
                output.WriteByte(TYPE_STRING);
                StringUtil.WriteUnicodeString(output, val);
                return;
            }
            if (value is ErrorConstant)
            {
                ErrorConstant ecVal = (ErrorConstant)value;
                output.WriteByte(TYPE_ERROR_CODE);
                long longVal = ecVal.ErrorCode;
                output.WriteLong(longVal);
                return;
            }

            throw new SystemException("Unexpected value type (" + value.GetType().Name + "'");
        }
    }
}
