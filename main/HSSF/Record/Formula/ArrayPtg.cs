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

namespace NPOI.HSSF.Record.Formula
{

    using System;
    using System.Text;
    using NPOI.Util;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Constant;
    
    using NPOI.Util.IO;

    /**
     * ArrayPtg - handles arrays
     * 
     * The ArrayPtg is a little weird, the size of the Ptg when parsing initially only
     * includes the Ptg sid and the reserved bytes. The next Ptg in the expression then follows.
     * It is only after the "size" of all the Ptgs is met, that the ArrayPtg data is actually
     * held after this. So Ptg.CreateParsedExpression keeps track of the number of 
     * ArrayPtg elements and need to Parse the data upto the FORMULA record size.
     *  
     * @author Jason Height (jheight at chariot dot net dot au)
     */
    public class ArrayPtg : Ptg
    {
        public const byte sid = 0x20;

        private const int RESERVED_FIELD_LEN = 7;
        	/** 
	 * The size of the plain tArray token written within the standard formula tokens
	 * (not including the data which comes after all formula tokens)
	 */
	public const int PLAIN_TOKEN_SIZE = 1+RESERVED_FIELD_LEN;

	private static byte[] DEFAULT_RESERVED_DATA = new byte[RESERVED_FIELD_LEN];

        // TODO - fix up field visibility and subclasses
        private byte[] field_1_reserved;

        // data from these fields comes after the Ptg data of all tokens in current formula
        private short token_1_columns;
        private short token_2_rows;
        private Object[] token_3_arrayValues;

        public ArrayPtg(LittleEndianInput in1)
        {
            field_1_reserved = new byte[RESERVED_FIELD_LEN];
            // TODO - Add ReadFully method to RecordInputStream
            for (int i = 0; i < RESERVED_FIELD_LEN; i++)
            {
                field_1_reserved[i] = (byte)in1.ReadByte();
            }
        }
        /**
         * @param values2d array values arranged in rows
         */
        public ArrayPtg(Object[][] values2d)
        {
            int nColumns = values2d[0].Length;
            int nRows = values2d.Length;
            // convert 2-d to 1-d array (row by row according to getValueIndex())
            token_1_columns = (short)nColumns;
            token_2_rows = (short)nRows;

            Object[] vv = new Object[token_1_columns * token_2_rows];
            for (int r = 0; r < nRows; r++)
            {
                Object[] rowData = values2d[r];
                for (int c = 0; c < nColumns; c++)
                {
                    vv[GetValueIndex(c, r)] = rowData[c];
                }
            }

            token_3_arrayValues = vv;
            field_1_reserved = DEFAULT_RESERVED_DATA;
        }
        public Object[][] GetTokenArrayValues()
        {
            if (token_3_arrayValues == null)
            {
                throw new InvalidOperationException("array values not read yet");
            }
            Object[][] result = new Object[token_2_rows][];
            for (int r = 0; r < token_2_rows; r++)
            {
                result[r]= new object[token_1_columns];
                for (int c = 0; c < token_1_columns; c++)
                {
                    result[r][c] = token_3_arrayValues[GetValueIndex(c, r)];
                }
            }
            return result;

        }

        public override bool IsBaseToken
        {
            get { return false; }
        }

        /** 
         * Read in the actual token (array) values. This occurs 
         * AFTER the last Ptg in the expression.
         * See page 304-305 of Excel97-2007BinaryFileFormat(xls)Specification.pdf
         */
        public void ReadTokenValues(LittleEndianInput in1)
        {
            short nColumns = (short)in1.ReadUByte();
            short nRows = in1.ReadShort();
            //The token_1_columns and token_2_rows do not follow the documentation.
            //The number of physical rows and columns is actually +1 of these values.
            //Which is not explicitly documented.
            nColumns++;
            nRows++;

            token_1_columns = nColumns;
            token_2_rows = nRows;

            int totalCount = nRows * nColumns;
            token_3_arrayValues = ConstantValueParser.Parse(in1, totalCount);
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder("[ArrayPtg]\n");

            buffer.Append("columns = ").Append(ColumnCount).Append("\n");
            buffer.Append("rows = ").Append(RowCount).Append("\n");
            for (int x = 0; x < ColumnCount; x++)
            {
                for (int y = 0; y < RowCount; y++)
                {
                    Object o = token_3_arrayValues.GetValue(GetValueIndex(x, y));
                    buffer.Append("[").Append(x).Append("][").Append(y).Append("] = ").Append(o).Append("\n");
                }
            }
            return buffer.ToString();
        }

        /**
         * Note - (2D) array elements are stored column by column 
         * @return the index into the internal 1D array for the specified column and row
         */
        /* package */
        public int GetValueIndex(int colIx, int rowIx)
        {
            if (colIx < 0 || colIx >= token_1_columns)
            {
                throw new ArgumentException("Specified colIx (" + colIx
                        + ") is outside the allowed range (0.." + (token_1_columns - 1) + ")");
            }
            if (rowIx < 0 || rowIx >= token_2_rows)
            {
                throw new ArgumentException("Specified rowIx (" + rowIx
                        + ") is outside the allowed range (0.." + (token_2_rows - 1) + ")");
            }
            return rowIx * token_1_columns + colIx;
        }

        public override void Write(LittleEndianOutput out1)
        {
       		out1.WriteByte(sid + PtgClass);
     		out1.Write(field_1_reserved);
        }
        public override void WriteBytes(byte[] data, int offset)
        {

            LittleEndian.PutByte(data, offset + 0, sid + PtgClass);
            Array.Copy(field_1_reserved, 0, data, offset + 1, RESERVED_FIELD_LEN);
        }

        public int WriteTokenValueBytes(LittleEndianOutput out1)
        {

		    out1.WriteByte(token_1_columns-1);
		    out1.WriteShort(token_2_rows-1);
            ConstantValueParser.Encode(out1, token_3_arrayValues);
            return 3 + ConstantValueParser.GetEncodedSize(token_3_arrayValues);
        }

        public short RowCount
        {
            get
            {
                return token_2_rows;
            }
        }

        public short ColumnCount
        {
            get
            {
                return token_1_columns;
            }
        }

        /** This size includes the size of the array Ptg plus the Array Ptg Token value size*/
        public override int Size
        {
            get
            {
                int size = 1 + 7 + 1 + 2;
                size += ConstantValueParser.GetEncodedSize(token_3_arrayValues);
                return size;
            }
        }

        public override String ToFormulaString()
        {
            StringBuilder b = new StringBuilder();
            b.Append("{");
            for (int x = 0; x < ColumnCount; x++)
            {
                if (x > 0)
                {
                    b.Append(";");
                }
                for (int y = 0; y < RowCount; y++)
                {
                    if (y > 0)
                    {
                        b.Append(",");
                    }
                    Object o = token_3_arrayValues.GetValue(GetValueIndex(x, y));
                    b.Append(GetConstantText(o));
                }
            }
            b.Append("}");
            return b.ToString();
        }

        private static String GetConstantText(Object o)
        {

            if (o == null)
            {
                return ""; // TODO - how is 'empty value' represented in formulas?
            }
            if (o is String)
            {
                return "\"" + (String)o + "\"";
            }
            if (o is Double)
            {
                return ((Double)o).ToString();
            }
            if (o is bool || o is Boolean)
            {
                return ((bool)o).ToString();
            }
            if (o is ErrorConstant)
            {
                return ((ErrorConstant)o).Text;
            }
            throw new ArgumentException("Unexpected constant class (" + o.GetType().Name + ")");
        }

        public override byte DefaultOperandClass
        {
            get { return Ptg.CLASS_ARRAY; }
        }

        public override Object Clone()
        {
            ArrayPtg ptg = (ArrayPtg)base.Clone();
            ptg.field_1_reserved = (byte[])field_1_reserved.Clone();
            ptg.token_3_arrayValues = (Object[])token_3_arrayValues.Clone();
            return ptg;
        }
    }
}