using System;
using System.Text;
using NPOI.Util;


namespace NPOI.HSSF.Record
{
    public class FtCblsSubRecord : SubRecord
    {
        public const short sid = 0x0C;
        private const int ENCODED_SIZE = 20;

        private byte[] reserved;

        /**
         * Construct a new <code>FtCblsSubRecord</code> and
         * fill its data with the default values
         */
        public FtCblsSubRecord()
        {
            reserved = new byte[ENCODED_SIZE];
        }

        public FtCblsSubRecord(ILittleEndianInput in1, int size)
        {
            if (size != ENCODED_SIZE)
            {
                throw new RecordFormatException("Unexpected size (" + size + ")");
            }
            //just grab the raw data
            byte[] buf = new byte[size];
            in1.ReadFully(buf);
            reserved = buf;
        }

        /**
         * Convert this record to string.
         * Used by BiffViewer and other utilities.
         */
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[FtCbls ]").Append("\n");
            buffer.Append("  size     = ").Append(DataSize).Append("\n");
            buffer.Append("  reserved = ").Append(HexDump.ToHex(reserved)).Append("\n");
            buffer.Append("[/FtCbls ]").Append("\n");
            return buffer.ToString();
        }

        /**
         * Serialize the record data into the supplied array of bytes
         *
         * @param out the stream to serialize into
         */
        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(sid);
            out1.WriteShort(reserved.Length);
            out1.Write(reserved);
        }

        public override int DataSize
        {
            get
            {
                return reserved.Length;
            }
        }

        /**
         * @return id of this record.
         */
        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        public override Object Clone()
        {
            FtCblsSubRecord rec = new FtCblsSubRecord();
            byte[] recdata = new byte[reserved.Length];
            Array.Copy(reserved, 0, recdata, 0, recdata.Length);
            rec.reserved = recdata;
            return rec;
        }

    }
}
