using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace NPOI.SS.Util
{
    using NPOI.Util;

    using NPOI.HSSF.Record;

    public class CellRangeAddressList
    {
        /**
 * List of <c>CellRangeAddress</c>es. Each structure represents a cell range
 */
        private readonly List<CellRangeAddress> _list;

        public CellRangeAddressList()
        {
            _list = new List<CellRangeAddress>();
        }
        /**
         * Convenience constructor for creating a <c>CellRangeAddressList</c> with a single 
         * <c>CellRangeAddress</c>.  Other <c>CellRangeAddress</c>es may be Added later.
         */
        public CellRangeAddressList(int firstRow, int lastRow, int firstCol, int lastCol)
            : this()
        {

            AddCellRangeAddress(firstRow, firstCol, lastRow, lastCol);
        }

        /**
         * @param in the RecordInputStream to read the record from
         */
        internal CellRangeAddressList(RecordInputStream in1)
        {
            int nItems = in1.ReadUShort();
            _list = new List<CellRangeAddress>(nItems);

            for (int k = 0; k < nItems; k++)
            {
                _list.Add(new CellRangeAddress(in1));
            }
        }

        public static CellRangeAddressList Parse(string cellRanges)
        {
            if(string.IsNullOrWhiteSpace(cellRanges))
            {
                throw new ArgumentException("cell range cannot be null or empty");
            }
            var ranges = cellRanges.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            var list = new CellRangeAddressList();
            foreach(var range in ranges)
            {
                var ca = CellRangeAddress.ValueOf(range.Trim());
                list.AddCellRangeAddress(ca);
            }
            return list;
        }

        /**
         * Get the number of following ADDR structures. The number of these
         * structures is automatically set when reading an Excel file and/or
         * increased when you manually Add a new ADDR structure . This is the reason
         * there isn't a set method for this field .
         *
         * @return number of ADDR structures
         */
        public int CountRanges()
        {
            return _list.Count;
        }

        public int NumberOfCells()
        {
            return _list.Sum(cr => cr.NumberOfCells);
        }

        /**
         * Add a cell range structure.
         *
         * @param firstRow - the upper left hand corner's row
         * @param firstCol - the upper left hand corner's col
         * @param lastRow - the lower right hand corner's row
         * @param lastCol - the lower right hand corner's col
         * @return the index of this ADDR structure
         */
        public void AddCellRangeAddress(int firstRow, int firstCol, int lastRow, int lastCol)
        {
            CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            AddCellRangeAddress(region);
        }
        public void AddCellRangeAddress(CellRangeAddress cra)
        {
            _list.Add(cra);
        }
        public CellRangeAddress Remove(int rangeIndex)
        {
            if (_list.Count == 0)
            {
                throw new Exception("List is empty");
            }
            if (rangeIndex < 0 || rangeIndex >= _list.Count)
            {
                throw new Exception("Range index (" + rangeIndex
                        + ") is outside allowable range (0.." + (_list.Count - 1) + ")");
            }
            CellRangeAddress cra = (CellRangeAddress)_list[rangeIndex];
            _list.Remove(_list[rangeIndex]);
            return cra;
        }

        /**
         * @return <c>CellRangeAddress</c> at the given index
         */
        public CellRangeAddress GetCellRangeAddress(int index)
        {
            return (CellRangeAddress)_list[index];
        }
        internal int Serialize(int offset, byte[] data)
        {
            int totalSize = this.Size;
            Serialize(new LittleEndianByteArrayOutputStream(data, offset, totalSize));
            return totalSize;
        }
        internal void Serialize(ILittleEndianOutput out1)
        {
            int nItems = _list.Count;
            out1.WriteShort(nItems);
            for (int k = 0; k < nItems; k++)
            {
                CellRangeAddress region = (CellRangeAddress)_list[k];
                region?.Serialize(out1);
            }
        }

        internal int Size
        {
            get
            {
                return GetEncodedSize(_list.Count);
            }
        }
        /**
         * @return the total size of for the specified number of ranges,
         *  including the initial 2 byte range count
         */
        internal static int GetEncodedSize(int numberOfRanges)
        {
            return 2 + CellRangeAddress.GetEncodedSize(numberOfRanges);
        }
        public CellRangeAddressList Copy()
        {
            CellRangeAddressList result = new CellRangeAddressList();

            int nItems = _list.Count;
            for (int k = 0; k < nItems; k++)
            {
                CellRangeAddress region = (CellRangeAddress)_list[k];
                if (region != null)
                    result.AddCellRangeAddress(region.Copy());
            }
            return result;
        }
        public CellRangeAddress[] CellRangeAddresses => _list.ToArray();
    }
}