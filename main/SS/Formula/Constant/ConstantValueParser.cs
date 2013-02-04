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

namespace NPOI.SS.Formula.Constant
{
    using System;
    using NPOI.Util;

    /**
     * To support Constant Values (2.5.7) as required by the CRN record.
     * This class is also used for two dimensional arrays which are encoded by 
     * EXTERNALNAME (5.39) records and Array tokens.<p/>
     * 
     * @author Josh Micich
     */
    public class ConstantValueParser
    {
        // note - these (non-combinable) enum values are sParse.
        private const int TYPE_EMPTY = 0;
        private const int TYPE_NUMBER = 1;
        private const int TYPE_STRING = 2;
        private const int TYPE_BOOLEAN = 4;
        private const int TYPE_ERROR_CODE = 16; // TODO - update OOO document to include this value

        private const int TRUE_ENCODING = 1;
        private const int FALSE_ENCODING = 0;

        // TODO - is this the best way to represent 'EMPTY'?
        private const object EMPTY_REPRESENTATION = null;

        private ConstantValueParser()
        {
            // no instances of this class
        }

        public static object[] Parse(ILittleEndianInput in1, int nValues)
        {
            object[] result = new Object[nValues];
            for (int i = 0; i < result.Length; i++)
            {
                result[i]=ReadAConstantValue(in1);
            }
            return result;
        }

        private static object ReadAConstantValue(ILittleEndianInput in1)
        {
            byte grbit = (byte)in1.ReadByte();
            switch (grbit)
            {
                case TYPE_EMPTY:
                    in1.ReadLong(); // 8 byte 'not used' field
                    return EMPTY_REPRESENTATION;
                case TYPE_NUMBER:
                    return in1.ReadDouble();
                case TYPE_STRING:
                    return StringUtil.ReadUnicodeString(in1);
                case TYPE_BOOLEAN:
                    return ReadBoolean(in1);
                case TYPE_ERROR_CODE:
                    int errCode = in1.ReadUShort();
                    // next 6 bytes are Unused
                    in1.ReadUShort();
                    in1.ReadInt();
                    return ErrorConstant.ValueOf(errCode);
            }
            throw new Exception("Unknown grbit value (" + grbit + ")");
        }

        private static Object ReadBoolean(ILittleEndianInput in1)
        {
            byte val = (byte)in1.ReadLong(); // 7 bytes 'not used'
            switch (val)
            {
                case FALSE_ENCODING:
                    return false;
                case TRUE_ENCODING:
                    return true;
            }
            // Don't tolerate Unusual bool encoded values (unless it becomes evident that they occur)
            throw new Exception("unexpected bool encoding (" + val + ")");
        }

        public static int GetEncodedSize(Array values)
        {
            // start with one byte 'type' code for each value
            int result = values.Length * 1;
            for (int i = 0; i < values.Length; i++)
            {
                result += GetEncodedSize(values.GetValue(i));
            }
            return result;
        }

        /**
         * @return encoded size without the 'type' code byte
         */
        private static int GetEncodedSize(Object obj)
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
            String strVal = (String)obj;
            return StringUtil.GetEncodedSize(strVal);
        }

        public static void Encode(ILittleEndianOutput out1, Array values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                EncodeSingleValue(out1, values.GetValue(i));
            }
        }

        private static void EncodeSingleValue(ILittleEndianOutput out1, Object value)
        {
            if (value == EMPTY_REPRESENTATION)
            {
                out1.WriteByte(TYPE_EMPTY);
                out1.WriteLong(0L);
                return;
            }
            if (value is bool)
            {
                bool bVal = ((bool)value);
                out1.WriteByte(TYPE_BOOLEAN);
                long longVal = bVal ? 1L : 0L;
                out1.WriteLong(longVal);
                return;
            }
            if (value is double)
            {
                double dVal = (double)value;
                out1.WriteByte(TYPE_NUMBER);
                out1.WriteDouble(dVal);
                return;
            }
            if (value is String)
            {
                String val = (String)value;
                out1.WriteByte(TYPE_STRING);
                StringUtil.WriteUnicodeString(out1, val);
                return;
            }
            if (value is ErrorConstant)
            {
                ErrorConstant ecVal = (ErrorConstant)value;
                out1.WriteByte(TYPE_ERROR_CODE);
                long longVal = ecVal.ErrorCode;
                out1.WriteLong(longVal);
                return;
            }

            throw new Exception("Unexpected value type (" + value.GetType().Name + "'");
        }
    }
}