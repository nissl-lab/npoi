using System;
using System.Text;
using System.IO;
using NPOI.Util;

namespace NPOI.HSSF.Util
{
    public class GUID {
        /*
         * this class is currently only used here, but could be moved to a
         * common package if needed
         */
        private const int TEXT_FORMAT_LENGTH = 36;

        public const int ENCODED_SIZE = 16;

        /** 4 bytes - little endian */
        private int _d1;
        /** 2 bytes - little endian */
        private int _d2;
        /** 2 bytes - little endian */
        private int _d3;
        /**
         * 8 bytes - serialized as big endian,  stored with inverted endianness here
         */
        private long _d4;

        public GUID(ILittleEndianInput in1) 
            :this(in1.ReadInt(), in1.ReadUShort(), in1.ReadUShort(), in1.ReadLong())
        {
            
        }

        public GUID(int d1, int d2, int d3, long d4) {
            _d1 = d1;
            _d2 = d2;
            _d3 = d3;
            _d4 = d4;
        }

        public void Serialize(ILittleEndianOutput out1) {
            out1.WriteInt(_d1);
            out1.WriteShort(_d2);
            out1.WriteShort(_d3);
            out1.WriteLong(_d4);
        }

        
        public override bool Equals(Object obj) {
            if (!(obj is GUID)) return false;
            GUID other = (GUID) obj;
            return _d1 == other._d1 && _d2 == other._d2 
                && _d3 == other._d3 && _d4 == other._d4;
        }

        public override int GetHashCode ()
        {
            return _d1 ^ _d2 ^ _d3 ^ _d4.GetHashCode ();
        }

        public int D1
        {
            get{
            return _d1;
            }
        }

        public int D2
        {
            get{
            return _d2;
            }
        }

        public int D3
        {
            get{
            return _d3;
            }
        }

        public long D4
        {
            get{
                //return _d4;
                byte[] buf;

                using (MemoryStream ms = new MemoryStream(8))
                {
                    BinaryWriter bw = new BinaryWriter(ms);
                    bw.Write(_d4);
                    buf = ms.ToArray();
                    bw.Close();
                }
                Array.Reverse(buf);
                return new LittleEndianByteArrayInputStream(buf).ReadLong();
            }
        }

        public String FormatAsString() {

            StringBuilder sb = new StringBuilder(36);

            int PREFIX_LEN = "0x".Length;
            sb.Append(HexDump.IntToHex(_d1), PREFIX_LEN, 8);
            sb.Append("-");
            sb.Append(HexDump.ShortToHex(_d2), PREFIX_LEN, 4);
            sb.Append("-");
            sb.Append(HexDump.ShortToHex(_d3), PREFIX_LEN, 4);
            sb.Append("-");
            char[] d4Chars = HexDump.LongToHex(D4);
            sb.Append(d4Chars, PREFIX_LEN, 4);
            sb.Append("-");
            sb.Append(d4Chars, PREFIX_LEN + 4, 12);
            return sb.ToString();
        }


        public override String ToString() {
            StringBuilder sb = new StringBuilder(64);
            sb.Append(GetType().Name).Append(" [");
            sb.Append(FormatAsString());
            sb.Append("]");
            return sb.ToString();
        }

        /**
         * Read a GUID in standard text form e.g.<br/>
         * 13579BDF-0246-8ACE-0123-456789ABCDEF 
         * <br/> -&gt; <br/>
         *  0x13579BDF, 0x0246, 0x8ACE 0x0123456789ABCDEF
         */
        public static GUID Parse(String rep) {
            char[] cc = rep.ToCharArray();
            if (cc.Length != TEXT_FORMAT_LENGTH) {
                throw new RecordFormatException("supplied text is the wrong length for a GUID");
            }
            int d0 = (ParseShort(cc, 0) << 16) + (ParseShort(cc, 4) << 0);
            int d1 = ParseShort(cc, 9);
            int d2 = ParseShort(cc, 14);
            for (int i = 23; i > 19; i--) {
                cc[i] = cc[i - 1];
            }
            long d3 = ParseLELong(cc, 20);

            return new GUID(d0, d1, d2, d3);
        }

        private static long ParseLELong(char[] cc, int startIndex) {
            long acc = 0;
            for (int i = startIndex + 14; i >= startIndex; i -= 2) {
                acc <<= 4;
                acc += ParseHexChar(cc[i + 0]);
                acc <<= 4;
                acc += ParseHexChar(cc[i + 1]);
            }
            return acc;
        }

        private static int ParseShort(char[] cc, int startIndex) {
            int acc = 0;
            for (int i = 0; i < 4; i++) {
                acc <<= 4;
                acc += ParseHexChar(cc[startIndex + i]);
            }
            return acc;
        }

        private static int ParseHexChar(char c) {
            if (c >= '0' && c <= '9') {
                return c - '0';
            }
            if (c >= 'A' && c <= 'F') {
                return c - 'A' + 10;
            }
            if (c >= 'a' && c <= 'f') {
                return c - 'a' + 10;
            }
            throw new RecordFormatException("Bad hex char '" + c + "'");
        }
    }

}
