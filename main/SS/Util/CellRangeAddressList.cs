using System;
using System.Collections;

namespace NPOI.SS.Util
{
    using NPOI.Util;

    using NPOI.HSSF.Record;

    public class CellRangeAddressList
    {
        /// <summary>
        /// List of <c>CellRangeAddress</c>es. Each structure represents a cell range
        /// </summary>
        private ArrayList _list;

        public CellRangeAddressList()
        {
            _list = new ArrayList();
        }
        /// <summary>
        /// Convenience constructor for creating a <c>CellRangeAddressList</c> with a single
        /// <c>CellRangeAddress</c>.  Other <c>CellRangeAddress</c>es may be Added later.
        /// </summary>
        public CellRangeAddressList(int firstRow, int lastRow, int firstCol, int lastCol)
            : this()
        {

            AddCellRangeAddress(firstRow, firstCol, lastRow, lastCol);
        }

        /// <summary>
        /// </summary>
        /// <param name="in">the RecordInputstream to read the record from</param>
        public CellRangeAddressList(RecordInputStream in1)
        {
            int nItems = in1.ReadUShort();
            _list = new ArrayList(nItems);

            for(int k = 0; k < nItems; k++)
            {
                _list.Add(new CellRangeAddress(in1));
            }
        }

        /// <summary>
        /// Get the number of following ADDR structures. The number of this
        /// structures is automatically set when reading an Excel file and/or
        /// increased when you manually Add a new ADDR structure . This is the reason
        /// there isn't a set method for this field .
        /// </summary>
        /// <return>number of ADDR structures</return>
        public int CountRanges()
        {
            return _list.Count;
        }

        /// <summary>
        /// Add a cell range structure.
        /// </summary>
        /// <param name="firstRow">- the upper left hand corner's row</param>
        /// <param name="firstCol">- the upper left hand corner's col</param>
        /// <param name="lastRow">- the lower right hand corner's row</param>
        /// <param name="lastCol">- the lower right hand corner's col</param>
        /// <return>the index of this ADDR structure</return>
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
            if(_list.Count == 0)
            {
                throw new Exception("List is empty");
            }
            if(rangeIndex < 0 || rangeIndex >= _list.Count)
            {
                throw new Exception("Range index (" + rangeIndex
                        + ") is outside allowable range (0.." + (_list.Count - 1) + ")");
            }
            CellRangeAddress cra = (CellRangeAddress)_list[rangeIndex];
            _list.Remove(rangeIndex);
            return cra;
        }

        /// <summary>
        /// </summary>
        /// <return><c>CellRangeAddress</c> at the given index</return>
        public CellRangeAddress GetCellRangeAddress(int index)
        {
            return (CellRangeAddress) _list[index];
        }
        public int Serialize(int offset, byte[] data)
        {
            int totalSize = this.Size;
            Serialize(new LittleEndianByteArrayOutputStream(data, offset, totalSize));
            return totalSize;
        }
        public void Serialize(ILittleEndianOutput out1)
        {
            int nItems = _list.Count;
            out1.WriteShort(nItems);
            for(int k = 0; k < nItems; k++)
            {
                CellRangeAddress region = (CellRangeAddress)_list[k];
                region.Serialize(out1);
            }
        }

        public int Size
        {
            get
            {
                return GetEncodedSize(_list.Count);
            }
        }
        /// <summary>
        /// </summary>
        /// <return>the total size of for the specified number of ranges,
        /// including the initial 2 byte range count
        /// </return>
        public static int GetEncodedSize(int numberOfRanges)
        {
            return 2 + CellRangeAddress.GetEncodedSize(numberOfRanges);
        }
        public CellRangeAddressList Copy()
        {
            CellRangeAddressList result = new CellRangeAddressList();

            int nItems = _list.Count;
            for(int k = 0; k < nItems; k++)
            {
                CellRangeAddress region = (CellRangeAddress)_list[k];
                result.AddCellRangeAddress(region.Copy());
            }
            return result;
        }
        public CellRangeAddress[] CellRangeAddresses
        {
            get
            {
                CellRangeAddress[] result =
                    (CellRangeAddress[])_list.ToArray(typeof(CellRangeAddress));
                return result;
            }
        }
    }
}