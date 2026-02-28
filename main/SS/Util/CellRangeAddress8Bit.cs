using System;

namespace NPOI.SS.Util
{
    using NPOI.Util;
    using NPOI.HSSF.Record;

    public class CellRangeAddress8Bit : CellRangeAddressBase
    {
        public const int ENCODED_SIZE = 6;

        public CellRangeAddress8Bit(int firstRow, int lastRow, int firstCol, int lastCol)
            : base(firstRow, lastRow, firstCol, lastCol)
        {

        }

        public CellRangeAddress8Bit(RecordInputStream in1)
            : base(ReadUShortAndCheck(in1), in1.ReadUShort(), in1.ReadUByte(), in1.ReadUByte())
        {

        }

        private static int ReadUShortAndCheck(RecordInputStream in1)
        {
            if (in1.Remaining < ENCODED_SIZE)
            {
                // Ran out of data
                throw new Exception("Ran out of data reading CellRangeAddress");
            }
            return in1.ReadUShort();
        }

        public int Serialize(int offset, byte[] data)
        {
            LittleEndian.PutUShort(data, offset + 0, FirstRow);
            LittleEndian.PutUShort(data, offset + 2, LastRow);
            LittleEndian.PutByte(data, offset + 4, FirstColumn);
            LittleEndian.PutByte(data, offset + 5, LastColumn);
            return ENCODED_SIZE;
        }
        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(FirstRow);
            out1.WriteShort(LastRow);
            out1.WriteByte(FirstColumn);
            out1.WriteByte(LastColumn);
        }
        public CellRangeAddress8Bit Copy()
        {
            return new CellRangeAddress8Bit(FirstRow, LastRow, FirstColumn, LastColumn);
        }

        public static int GetEncodedSize(int numberOfItems)
        {
            return numberOfItems * ENCODED_SIZE;
        }
    }
}