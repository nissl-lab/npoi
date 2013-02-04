
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
        

/*
 * BoolErrRecord.java
 *
 * Created on January 19, 2002, 9:30 AM
 */
namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using NPOI.Util;
    using NPOI.SS.UserModel;

    /**
     * Creates new BoolErrRecord. 
     * REFERENCE:  PG ??? Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Michael P. Harhen
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class BoolErrRecord : CellRecord
    {
        public const short sid = 0x205;
        private int _value;
        /**
	 * If <code>true</code>, this record represents an error cell value, otherwise this record represents a boolean cell value
	 */
        private bool _isError;

        /** Creates new BoolErrRecord */

        public BoolErrRecord()
        {
            // fields uninitialised
        }

        /**
         * Constructs a BoolErr record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public BoolErrRecord(RecordInputStream in1)
            : base(in1)
        {
            switch (in1.Remaining)
            {
                case 2:
                    _value = in1.ReadByte();
                    break;
                case 3:
                    _value = in1.ReadUShort();
                    break;
                default:
                    throw new RecordFormatException("Unexpected size ("
                            + in1.Remaining + ") for BOOLERR record.");
            }
            int flag = in1.ReadUByte();
            switch (flag)
            {
                case 0:
                    _isError = false;
                    break;
                case 1:
                    _isError = true;
                    break;
                default:
                    throw new RecordFormatException("Unexpected isError flag ("
                            + flag + ") for BOOLERR record.");
            }
        }


        /**
         * Set the bool value for the cell
         *
         * @param value   representing the bool value
         */

        public void SetValue(bool value)
        {
            _value = value ? 1 : 0;
            _isError = false;
        }

        /**
         * Set the error value for the cell
         *
         * @param value     error representing the error value
         *                  this value can only be 0,7,15,23,29,36 or 42
         *                  see bugzilla bug 16560 for an explanation
         */

        public void SetValue(byte value)
        {
            switch (value)
            {
                case ErrorConstants.ERROR_NULL:
                case ErrorConstants.ERROR_DIV_0:
                case ErrorConstants.ERROR_VALUE:
                case ErrorConstants.ERROR_REF:
                case ErrorConstants.ERROR_NAME:
                case ErrorConstants.ERROR_NUM:
                case ErrorConstants.ERROR_NA:
                    _value = value;
                    _isError = true;
                    return;
            }
            throw new ArgumentException("Error Value can only be 0,7,15,23,29,36 or 42. It cannot be " + value);
        }

        /**
         * Get the value for the cell
         *
         * @return bool representing the bool value
         */

        public bool BooleanValue
        {
            get { return (_value != 0); }
        }

        /**
         * Get the error value for the cell
         *
         * @return byte representing the error value
         */

        public byte ErrorValue
        {
            get { return (byte)_value; }
        }
        /**
     * Indicates whether the call holds a boolean value
     *
     * @return boolean true if the cell holds a boolean value
     */

        public bool IsBoolean
        {
            get { return !_isError; }
        }

        /**
         * Indicates whether the call holds an error value
         *
         * @return bool true if the cell holds an error value
         */

        public bool IsError
        {
            get { return _isError; }
        }

        protected override String RecordName
        {
            get { return "BOOLERR"; }
        }
        protected override void AppendValueText(StringBuilder buffer)
        {
            if (IsBoolean)
            {
                buffer.Append("    .boolValue   = ").Append(BooleanValue)
                    .Append("\n");
            }
            else
            {
                buffer.Append("    .errCode     = ").Append(ErrorValue)
                    .Append("\n");
            }
        }

        protected override void SerializeValue(ILittleEndianOutput out1)
        {
            out1.WriteByte(_value);
            out1.WriteByte(_isError ? 1 : 0);
        }
        protected override int ValueDataSize
        {
            get
            {
                return 2;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }


        public override Object Clone()
        {
            BoolErrRecord rec = new BoolErrRecord();
            CopyBaseFields(rec);
            rec._value = _value;
            rec._isError = _isError;
            return rec;
        }
    }
}