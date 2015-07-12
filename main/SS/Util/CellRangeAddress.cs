using System;
using System.Text;
using NPOI.Util;
using NPOI.HSSF.Record;

using NPOI.SS.Formula;

namespace NPOI.SS.Util
{
    public class CellRangeAddress : CellRangeAddressBase
    {
        public const int ENCODED_SIZE = 8;
        /**
         * Creates new cell range. Indexes are zero-based.
         * 
         * @param firstRow Index of first row
         * @param lastRow Index of last row (inclusive), must be equal to or larger than {@code firstRow}
         * @param firstCol Index of first column
         * @param lastCol Index of last column (inclusive), must be equal to or larger than {@code firstCol}
         */
        public CellRangeAddress(int firstRow, int lastRow, int firstCol, int lastCol)
            : base(firstRow, lastRow, firstCol, lastCol)
        {
            if (lastRow < firstRow || lastCol < firstCol)
                throw new ArgumentException("lastRow < firstRow || lastCol < firstCol");
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
                throw new RuntimeException("Ran out of data readin1g CellRangeAddress");
            }
            return in1.ReadUShort();
        }

        public void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(FirstRow);
            out1.WriteShort(LastRow);
            out1.WriteShort(FirstColumn);
            out1.WriteShort(LastColumn);
        }
        public String FormatAsString()
        {
            return FormatAsString(null, false);
        }
        /**
         * @return the text format of this range using specified sheet name.
         */
        public String FormatAsString(String sheetName, bool useAbsoluteAddress)
        {
            StringBuilder sb = new StringBuilder();
            if (sheetName != null)
            {
                sb.Append(SheetNameFormatter.Format(sheetName));
                sb.Append("!");
            }
            CellReference cellRefFrom = new CellReference(FirstRow, FirstColumn,
                    useAbsoluteAddress, useAbsoluteAddress);
            CellReference cellRefTo = new CellReference(LastRow, LastColumn,
                    useAbsoluteAddress, useAbsoluteAddress);
            sb.Append(cellRefFrom.FormatAsString());

            //for a single-cell reference return A1 instead of A1:A1
            //for full-column ranges or full-row ranges return A:A instead of A,
            //and 1:1 instead of 1
            if (!cellRefFrom.Equals(cellRefTo)
                || IsFullColumnRange || IsFullRowRange)
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

        /// <summary>
        /// Creates a CellRangeAddress from a cell range reference string.
        /// </summary>
        /// <param name="reference">
        /// usually a standard area ref (e.g. "B1:D8").  May be a single 
        /// cell ref (e.g. "B5") in which case the result is a 1 x 1 cell 
        /// range. May also be a whole row range (e.g. "3:5"), or a whole 
        /// column range (e.g. "C:F")
        /// </param>
        /// <returns>a CellRangeAddress object</returns>
        public static CellRangeAddress ValueOf(String reference)
        {
            int sep = reference.IndexOf(":", StringComparison.Ordinal);
            CellReference a;
            CellReference b;
            if (sep == -1)
            {
                a = new CellReference(reference);
                b = a;
            }
            else
            {
                a = new CellReference(reference.Substring(0, sep));
                b = new CellReference(reference.Substring(sep + 1));
            }
            return new CellRangeAddress(a.Row, b.Row, a.Col, b.Col);
        }
    }
}
