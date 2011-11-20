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

namespace NPOI.SS.Formula.PTG
{
using NPOI.SS.Formula.Constant;
using NPOI.SS.Util;
using NPOI.Util;
using System;
using System.Text;
    using NPOI.Util.IO;

/**
 * ArrayPtg - handles arrays
 *
 * The ArrayPtg is a little weird, the size of the Ptg when parsing initially only
 * includes the Ptg sid and the reserved bytes. The next Ptg in the expression then follows.
 * It is only after the "size" of all the Ptgs is met, that the ArrayPtg data is actually
 * held after this. So Ptg.createParsedExpression keeps track of the number of
 * ArrayPtg elements and need to parse the data upto the FORMULA record size.
 *
 * @author Jason Height (jheight at chariot dot net dot au)
 */
public  class ArrayPtg : Ptg {
	public const  byte Sid  = 0x20;

	private static  int RESERVED_FIELD_LEN = 7;
	/**
	 * The size of the plain tArray token written within the standard formula tokens
	 * (not including the data which comes after all formula tokens)
	 */
	public static  int PLAIN_TOKEN_SIZE = 1+RESERVED_FIELD_LEN;

	// 7 bytes of data (stored as an int, short and byte here)
	private  int _reserved0Int;
	private  int _reserved1Short;
	private  int _reserved2Byte;

	// data from these fields comes after the Ptg data of all tokens in current formula
	private  int  _nColumns;
	private  int _nRows;
	private  object[] _arrayValues;

	ArrayPtg(int reserved0, int reserved1, int reserved2, int nColumns, int nRows, Object[] arrayValues) {
		_reserved0Int = reserved0;
		_reserved1Short = reserved1;
		_reserved2Byte = reserved2;
		_nColumns = nColumns;
		_nRows = nRows;
		_arrayValues = arrayValues;
	}
	/**
	 * @param values2d array values arranged in rows
	 */
	public ArrayPtg(object[][] values2d) {
		int nColumns = values2d[0].Length;
		int nRows = values2d.Length;
		// convert 2-d to 1-d array (row by row according to getValueIndex())
		_nColumns = (short) nColumns;
		_nRows = (short) nRows;

		object[] vv = new object[_nColumns * _nRows];
		for (int r=0; r<nRows; r++) {
			object[] rowData = values2d[r];
			for (int c=0; c<nColumns; c++) {
				vv[GetValueIndex(c, r)] = rowData[c];
			}
		}

		_arrayValues = vv;
		_reserved0Int = 0;
		_reserved1Short = 0;
		_reserved2Byte = 0;
	}
	/**
	 * @return 2-d array (inner index is rowIx, outer index is colIx)
	 */
	public Object[][] GetTokenArrayValues() {
		if (_arrayValues == null) {
			throw new ApplicationException("array values not read yet");
		}
		Object[][] result = new Object[_nRows][];
		for (int r = 0; r < _nRows; r++) {
			//Object[] rowData = result[r];
            Object[] rowData = new Object[_nColumns];
			for (int c = 0; c < _nColumns; c++) {
				rowData[c] = _arrayValues[GetValueIndex(c, r)];
			}
		}
		return result;
	}

    public override bool IsBaseToken()
    {
		return false;
	}

    public override String ToString()
    {
		StringBuilder sb = new StringBuilder("[ArrayPtg]\n");

		sb.Append("nRows = ").Append(RowCount).Append("\n");
		sb.Append("nCols = ").Append(ColumnCount).Append("\n");
		if (_arrayValues == null) {
			sb.Append("  #values#uninitialised#\n");
		} else {
			sb.Append("  ").Append(ToFormulaString());
		}
		return sb.ToString();
	}

	/**
	 * Note - (2D) array elements are stored row by row
	 * @return the index into the internal 1D array for the specified column and row
	 */
	/* package */ int GetValueIndex(int colIx, int rowIx) {
		if(colIx < 0 || colIx >= _nColumns) {
			throw new ArgumentException("Specified colIx (" + colIx
					+ ") is outside the allowed range (0.." + (_nColumns-1) + ")");
		}
		if(rowIx < 0 || rowIx >= _nRows) {
			throw new ArgumentException("Specified rowIx (" + rowIx
					+ ") is outside the allowed range (0.." + (_nRows-1) + ")");
		}
		return rowIx * _nColumns + colIx;
	}

    public override void Write(LittleEndianOutput output)
    {
		output.WriteByte(Sid + GetPtgClass());
		output.WriteInt(_reserved0Int);
		output.WriteShort(_reserved1Short);
		output.WriteByte(_reserved2Byte);
	}

	public int WriteTokenValueBytes(LittleEndianOutput output) {

		output.WriteByte(_nColumns-1);
		output.WriteShort(_nRows-1);
		ConstantValueParser.Encode(output, _arrayValues);
		return 3 + ConstantValueParser.GetEncodedSize(_arrayValues);
	}

	public int RowCount {
		get
        {
        return _nRows;
        }
	}

	public int ColumnCount {
        get{
		return _nColumns;
        }
	}

	/** This size includes the size of the array Ptg plus the Array Ptg Token value size*/
    public override int GetSize()
    {
		return PLAIN_TOKEN_SIZE
			// data written after the all tokens:
			+ 1 + 2 // column, row
			+ ConstantValueParser.GetEncodedSize(_arrayValues);
	}

    public override String ToFormulaString()
    {
		StringBuilder b = new StringBuilder();
		b.Append("{");
	  	for (int y=0;y<RowCount;y++) {
			if (y > 0) {
				b.Append(";");
			}
			for (int x=0;x<ColumnCount;x++) {
			  	if (x > 0) {
					b.Append(",");
				}
		  		Object o = _arrayValues[GetValueIndex(x, y)];
		  		b.Append(GetConstantText(o));
		  	}
		  }
		b.Append("}");
		return b.ToString();
	}

	private static string GetConstantText(Object o) {

		if (o == null) {
			throw new ApplicationException("Array item cannot be null");
		}
		if (o is String) {
			return "\"" + (String)o + "\"";
		}
		if (o is Double) {
			return o.ToString();
		}
		if (o is Boolean) {
			return (Boolean)o ? "TRUE" : "FALSE";
		}
		if (o is ErrorConstant) {
			return ((ErrorConstant)o).GetText();
		}
		throw new ArgumentException("Unexpected constant class (" + o.GetType().ToString() + ")");
	}

    public override byte GetDefaultOperandClass()
    {
		return Ptg.CLASS_ARRAY;
	}

	/**
	 * Represents the initial plain tArray token (without the constant data that trails the whole
	 * formula).  Objects of this class are only temporary and cannot be used as {@link Ptg}s.
	 * These temporary objects get converted to {@link ArrayPtg} by the
	 * {@link #finishReading(LittleEndianInput)} method.
	 */
	public class Initial : Ptg {
		private  int _reserved0;
		private  int _reserved1;
		private  int _reserved2;

		public Initial(LittleEndianInput input) {
			_reserved0 = input.ReadInt();
			_reserved1 = input.ReadUShort();
			_reserved2 = input.ReadUByte();
		}
		private static void Invalid() {
			throw new ApplicationException("This object is a partially initialised tArray, and cannot be used as a Ptg");
		}
        public override byte GetDefaultOperandClass()
        {
			//throw Invalid();
            throw new NotImplementedException();
		}
		public override int GetSize() {
			return PLAIN_TOKEN_SIZE;
		}
        public override bool IsBaseToken()
        {
			return false;
		}
        public override String ToFormulaString()
        {
			//throw Invalid();
            throw new NotImplementedException();
		}
        public override void Write(LittleEndianOutput output)
        {
			//throw Invalid();
            throw new NotImplementedException();
		}
		/**
		 * Read in the actual token (array) values. This occurs
		 * AFTER the last Ptg in the expression.
		 * See page 304-305 of Excel97-2007BinaryFileFormat(xls)Specification.pdf
		 */
		public ArrayPtg FinishReading(LittleEndianInput input) {
			int nColumns = input.ReadUByte();
			short nRows = input.ReadShort();
			//The token_1_columns and token_2_rows do not follow the documentation.
			//The number of physical rows and columns is actually +1 of these values.
			//Which is not explicitly documented.
			nColumns++;
			nRows++;

			int totalCount = nRows * nColumns;
			Object[] arrayValues = ConstantValueParser.Parse(input, totalCount);

			ArrayPtg result = new ArrayPtg(_reserved0, _reserved1, _reserved2, nColumns, nRows, arrayValues);
			//result.SetClass(GetPtgClass());
			return result;
		}
	}
}
}
