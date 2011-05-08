using System;
using System.Collections.Generic;
using System.Text;
using NPOI.Util;
using NPOI.HSSF.Record;
using NPOI.Util.IO;

namespace NPOI.SS.Util
{
    public class CellRangeAddress : CellRangeAddressBase
    {
        public static int ENCODED_SIZE = 8;

        public CellRangeAddress(int firstRow, int lastRow, int firstCol, int lastCol)
            : base(firstRow, lastRow, firstCol, lastCol)
        {

        }

        public CellRangeAddress(RecordInputStream in1)
            : base(ReadUShortAndCheck(in1), in1.ReadUShort(), in1.ReadUShort(), in1.ReadUShort())
        {

        }

        private static int ReadUShortAndCheck(RecordInputStream in1)
        {
            if (in1.Remaining < ENCODED_SIZE)
            {
                // Ran out of data
                throw new Exception("Ran out of data readin1g CellRangeAddress");
            }
            return in1.ReadUShort();
        }

        public int Serialize(int offset, byte[] data)
        {
            LittleEndian.PutUShort(data, offset + 0, FirstRow);
            LittleEndian.PutUShort(data, offset + 2, LastRow);
            LittleEndian.PutUShort(data, offset + 4, FirstColumn);
            LittleEndian.PutUShort(data, offset + 6, LastColumn);
            return ENCODED_SIZE;
        }
        public void Serialize(LittleEndianOutput out1)
        {
            out1.WriteShort(FirstRow);
            out1.WriteShort(LastRow);
            out1.WriteShort(FirstColumn);
            out1.WriteShort(LastColumn);
        }
        public String FormatAsString()
        {
            StringBuilder sb = new StringBuilder();
            CellReference cellRefFrom = new CellReference(FirstRow, FirstColumn);
            CellReference cellRefTo = new CellReference(LastRow, LastColumn);
            sb.Append(cellRefFrom.FormatAsString());
            if (!cellRefFrom.Equals(cellRefTo))
            {
                sb.Append(':');
                sb.Append(cellRefTo.FormatAsString());
            }
            return sb.ToString();
        }
        public CellRangeAddress Copy()
        {
            return new CellRangeAddress(FirstRow, LastRow, FirstColumn, LastColumn);
        }

        public static int GetEncodedSize(int numberOfItems)
        {
            return numberOfItems * ENCODED_SIZE;
        }
    }
}
