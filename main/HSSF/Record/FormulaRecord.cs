
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional inFormation regarding copyright ownership.
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
 * FormulaRecord.java
 *
 * Created on October 28, 2001, 5:44 PM
 */

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;


    /**
 * Manages the cached formula result values of other types besides numeric.
 * Excel encodes the same 8 bytes that would be field_4_value with various NaN
 * values that are decoded/encoded by this class. 
 */
    internal class SpecialCachedValue
    {
        /** deliberately chosen by Excel in order to encode other values within Double NaNs */
        private const long BIT_MARKER = unchecked((long)0xFFFF000000000000L);
        private const int VARIABLE_DATA_LENGTH = 6;
        private const int DATA_INDEX = 2;

        public const int STRING = 0;
        public const int BOOLEAN = 1;
        public const int ERROR_CODE = 2;
        public const int EMPTY = 3;

        private byte[] _variableData;

        private SpecialCachedValue(byte[] data)
        {
            _variableData = data;
        }
        public int GetTypeCode()
        {
            return _variableData[0];
        }

        /**
         * @return <c>null</c> if the double value encoded by <c>valueLongBits</c> 
         * is a normal (non NaN) double value.
         */
        public static SpecialCachedValue Create(long valueLongBits)
        {
            if ((BIT_MARKER & valueLongBits) != BIT_MARKER)
            {
                return null;
            }

            byte[] result = new byte[VARIABLE_DATA_LENGTH];
            long x = valueLongBits;
            for (int i = 0; i < VARIABLE_DATA_LENGTH; i++)
            {
                result[i] = (byte)x;
                x >>= 8;
            }
            switch (result[0])
            {
                case STRING:
                case BOOLEAN:
                case ERROR_CODE:
                case EMPTY:
                    break;
                default:
                    throw new RecordFormatException("Bad special value code (" + result[0] + ")");
            }
            return new SpecialCachedValue(result);
        }
        //public void Serialize(byte[] data, int offset)
        //{
        //    System.Array.Copy(_variableData, 0, data, offset, VARIABLE_DATA_LENGTH);
        //    LittleEndian.PutUShort(data, offset + VARIABLE_DATA_LENGTH, 0xFFFF);
        //}
        public void Serialize(ILittleEndianOutput out1) {
            out1.Write(_variableData);
            out1.WriteShort(0xFFFF);
        }

        public String FormatDebugString
        {
            get
            {
                return FormatValue + ' ' + HexDump.ToHex(_variableData);
            }
        }
        private String FormatValue
        {
            get
            {
                int typeCode = GetTypeCode();
                switch (typeCode)
                {
                    case STRING: return "<string>";
                    case BOOLEAN: return DataValue == 0 ? "FALSE" : "TRUE";
                    case ERROR_CODE: return ErrorEval.GetText(DataValue);
                    case EMPTY: return "<empty>";
                }
                return "#error(type=" + typeCode + ")#";
            }
        }
        private int DataValue
        {
            get
            {
                return _variableData[DATA_INDEX];
            }
        }
        public static SpecialCachedValue CreateCachedEmptyValue()
        {
            return Create(EMPTY, 0);
        }
        public static SpecialCachedValue CreateForString()
        {
            return Create(STRING, 0);
        }
        public static SpecialCachedValue CreateCachedBoolean(bool b)
        {
            return Create(BOOLEAN, b ? 1 : 0);
        }
        public static SpecialCachedValue CreateCachedErrorCode(int errorCode)
        {
            return Create(ERROR_CODE, errorCode);
        }
        private static SpecialCachedValue Create(int code, int data)
        {
            byte[] vd = {
                    (byte) code,
                    0,
                    (byte) data,
                    0,
                    0,
                    0,
            };
            return new SpecialCachedValue(vd);
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name);
            sb.Append('[').Append(FormatValue).Append(']');
            return sb.ToString();
        }
        public NPOI.SS.UserModel.CellType GetValueType()
        {
            int typeCode = GetTypeCode();
            switch (typeCode)
            {
                case STRING: return NPOI.SS.UserModel.CellType.String;
                case BOOLEAN: return NPOI.SS.UserModel.CellType.Boolean;
                case ERROR_CODE: return NPOI.SS.UserModel.CellType.Error;
                case EMPTY: return NPOI.SS.UserModel.CellType.String; // is this correct?
            }
            throw new InvalidOperationException("Unexpected type id (" + typeCode + ")");
        }
        public bool GetBooleanValue()
        {
            if (GetTypeCode() != BOOLEAN)
            {
                throw new InvalidOperationException("Not a bool cached value - " + FormatValue);
            }
            return DataValue != 0;
        }
        public int GetErrorValue()
        {
            if (GetTypeCode() != ERROR_CODE)
            {
                throw new InvalidOperationException("Not an error cached value - " + FormatValue);
            }
            return DataValue;
        }
    }

    /**
     * Formula Record.
     * REFERENCE:  PG 317/444 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */
    [Serializable]
    public class FormulaRecord : CellRecord
    {

        public const short sid = 0x06;   // docs say 406...because of a bug Microsoft support site article #Q184647)
        private const int FIXED_SIZE = 14;


        private double field_4_value;
        private short field_5_options;
        private BitField alwaysCalc = BitFieldFactory.GetInstance(0x0001);
        private BitField calcOnLoad = BitFieldFactory.GetInstance(0x0002);
        private BitField sharedFormula = BitFieldFactory.GetInstance(0x0008);
        private int field_6_zero;
        [NonSerialized]
        private NPOI.SS.Formula.Formula field_8_parsed_expr;

        /**
        * Since the NaN support seems sketchy (different constants) we'll store and spit it out directly
        */
        [NonSerialized]
        private SpecialCachedValue specialCachedValue;


        /*
         * Since the NaN support seems sketchy (different constants) we'll store and spit it out directly
         */
        // fix warning CS0169 "never used": private byte[] value_data;
        // fix warning CS0169 "never used": private byte[] all_data; //if formula support is not enabled then
        //we'll just store/reSerialize

        /** Creates new FormulaRecord */

        public FormulaRecord()
        {
            field_8_parsed_expr = NPOI.SS.Formula.Formula.Create(Ptg.EMPTY_PTG_ARRAY);
        }

        /**
         * Constructs a Formula record and Sets its fields appropriately.
         * Note - id must be 0x06 (NOT 0x406 see MSKB #Q184647 for an 
         * "explanation of this bug in the documentation) or an exception
         *  will be throw upon validation
         *
         * @param in the RecordInputstream to Read the record from
         */

        public FormulaRecord(RecordInputStream in1):base(in1)
        {
                long valueLongBits  = in1.ReadLong();
                field_5_options = in1.ReadShort();
                specialCachedValue = SpecialCachedValue.Create(valueLongBits);
                if (specialCachedValue == null) {
                    field_4_value = BitConverter.Int64BitsToDouble(valueLongBits);
                }

                field_6_zero = in1.ReadInt();
                int field_7_expression_len = in1.ReadShort();

                field_8_parsed_expr = NPOI.SS.Formula.Formula.Read(field_7_expression_len, in1,in1.Available());
        }
        /**
 * @return <c>true</c> if this {@link FormulaRecord} is followed by a
 *  {@link StringRecord} representing the cached text result of the formula
 *  evaluation.
 */
        public bool HasCachedResultString
        {
            get
            {
                if (specialCachedValue == null)
                {
                    return false;
                }
                return specialCachedValue.GetTypeCode() == SpecialCachedValue.STRING;
            }
        }
        public void SetParsedExpression(Ptg[] ptgs)
        {
            field_8_parsed_expr = NPOI.SS.Formula.Formula.Create(ptgs);
        }
        public void SetSharedFormula(bool flag)
        {
            field_5_options =
                sharedFormula.SetShortBoolean(field_5_options, flag);
        }

        /**
         * Get the calculated value of the formula
         *
         * @return calculated value
         */
        public double Value
        {
            get { return field_4_value; }
            set { 
                field_4_value = value; 
                specialCachedValue = null;
            }
        }

        /**
         * Get the option flags
         *
         * @return bitmask
         */
        public short Options
        {
            get { return field_5_options; }
            set { field_5_options = value; }
        }

        public bool IsSharedFormula
        {
            get { return sharedFormula.IsSet(field_5_options); }
            set
            {
                field_5_options =
                    sharedFormula.SetShortBoolean(field_5_options, value);
            }
        }

        public bool IsAlwaysCalc
        {
            get { return alwaysCalc.IsSet(field_5_options); }
            set
            {
                field_5_options =
                    alwaysCalc.SetShortBoolean(field_5_options, value);
            }
        }

        public bool IsCalcOnLoad
        {
            get { return calcOnLoad.IsSet(field_5_options); }
            set
            {
                field_5_options =
                    calcOnLoad.SetShortBoolean(field_5_options, value);
            }
        }

        /**
         * Get the stack as a list
         *
         * @return list of tokens (casts stack to a list and returns it!)
         * this method can return null Is we are Unable to Create Ptgs from 
         *     existing excel file
         * callers should Check for null!
         */

        public Ptg[] ParsedExpression
        {
            get { return (Ptg[])field_8_parsed_expr.Tokens; }
            set { field_8_parsed_expr = NPOI.SS.Formula.Formula.Create(value); }
        }
        public NPOI.SS.Formula.Formula Formula
        {
            get
            {
                return field_8_parsed_expr;
            }
        }

        protected override String RecordName
        {
            get
            {
                return "FORMULA";
            }
        }

        protected override int ValueDataSize 
        {
            get
            {
                return FIXED_SIZE + field_8_parsed_expr.EncodedSize;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public void SetCachedResultTypeEmptyString()
        {
            specialCachedValue = SpecialCachedValue.CreateCachedEmptyValue();
        }
        public void SetCachedResultTypeString()
        {
            specialCachedValue = SpecialCachedValue.CreateForString();
        }
        public void SetCachedResultErrorCode(int errorCode)
        {
            specialCachedValue = SpecialCachedValue.CreateCachedErrorCode(errorCode);
        }
        public void SetCachedResultBoolean(bool value)
        {
            specialCachedValue = SpecialCachedValue.CreateCachedBoolean(value);
        }
        public bool CachedBooleanValue
        {
            get
            {
                return specialCachedValue.GetBooleanValue();
            }
        }
        public int CachedErrorValue
        {
            get
            {
                return specialCachedValue.GetErrorValue();
            }
        }

        public NPOI.SS.UserModel.CellType CachedResultType
        {
            get
            {
                if (specialCachedValue == null)
                {
                    return NPOI.SS.UserModel.CellType.Numeric;
                }
                return (NPOI.SS.UserModel.CellType)specialCachedValue.GetValueType();
            }
        }

        public override bool Equals(Object obj)
        {
            if (!(obj is CellValueRecordInterface))
            {
                return false;
            }
            CellValueRecordInterface loc = (CellValueRecordInterface)obj;

            if ((this.Row == loc.Row)
                    && (this.Column == loc.Column))
            {
                return true;
            }
            return false;
        }

        public override int GetHashCode ()
        {
            return Row ^ Column;
        }

        protected override void SerializeValue(ILittleEndianOutput out1)
        {
            if (specialCachedValue == null)
            {
                out1.WriteDouble(field_4_value);
            }
            else
            {
                specialCachedValue.Serialize(out1);
            }

            out1.WriteShort(Options);

            out1.WriteInt(field_6_zero); // may as well write original data back so as to minimise differences from original
            field_8_parsed_expr.Serialize(out1);
        }

        protected override void AppendValueText(StringBuilder buffer)
        {
            buffer.Append("    .value           = ");
            if (specialCachedValue == null)
            {
                buffer.Append(field_4_value).Append("\n");
            }
            else
            {
                buffer.Append(specialCachedValue.FormatDebugString).Append("\n");
            }
            buffer.Append("    .options         = ").Append(Options).Append("\n");
            buffer.Append("      .alwaysCalc         = ").Append(alwaysCalc.IsSet(Options)).Append("\n");
            buffer.Append("      .CalcOnLoad         = ").Append(calcOnLoad.IsSet(Options)).Append("\n");
            buffer.Append("      .sharedFormula         = ").Append(sharedFormula.IsSet(Options)).Append("\n");
            buffer.Append("    .zero            = ").Append(field_6_zero).Append("\n");

            Ptg[] ptgs = field_8_parsed_expr.Tokens;
            for (int k = 0; k < ptgs.Length; k++)
            {
                buffer.Append("	 Ptg[").Append(k).Append("]=");
                Ptg ptg = ptgs[k];
                buffer.Append(ptg.ToString()).Append(ptg.RVAType).Append("\n");
            }
        }

        public override Object Clone()
        {
            FormulaRecord rec = new FormulaRecord();
            CopyBaseFields(rec);
            rec.field_4_value = field_4_value;
            rec.field_5_options = field_5_options;
            rec.field_6_zero = field_6_zero;
            rec.field_8_parsed_expr = field_8_parsed_expr.Copy();
            rec.specialCachedValue = specialCachedValue;
            return rec;
        }

    }
}