
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

namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;
    using System.IO;

    using NPOI.Util;

    using NPOI.SS.Formula.PTG;
    using System.Globalization;

    /**
     * A sub-record within the OBJ record which stores a reference to an object
     * stored in a Separate entry within the OLE2 compound file.
     *
     * @author Daniel Noll
     */
    public class EmbeddedObjectRefSubRecord
       : SubRecord
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(EmbeddedObjectRefSubRecord));
        public const short sid = 0x9;
        private static byte[] EMPTY_BYTE_ARRAY = { };

        private int field_1_unknown_int;                            // Unknown stuff at the front.  TODO: Confirm that it's a short[]
        /** either an area or a cell ref */
        private Ptg field_2_refPtg;
        private byte[] field_2_unknownFormulaData;
        // TODO: Consider making a utility class for these.  I've discovered the same field ordering
        //       in FormatRecord and StringRecord, it may be elsewhere too.
        public bool field_3_unicode_flag;                        // Flags whether the string is Unicode.
        private String field_4_ole_classname; // Classname of the embedded OLE document (e.g. Word.Document.8)
        /** Formulas often have a single non-zero trailing byte.
         * This is in a similar position to he pre-streamId padding
         * It is unknown if the value is important (it seems to mirror a value a few bytes earlier) 
         *  */
        private Byte? field_4_unknownByte;
        private int? field_5_stream_id; // ID of the OLE stream containing the actual data.
        private byte[] field_6_unknown;

        public EmbeddedObjectRefSubRecord()
        {
            field_2_unknownFormulaData = new byte[] { 0x02, 0x6C, 0x6A, 0x16, 0x01, }; // just some sample data.  These values vary a lot
            field_6_unknown = EMPTY_BYTE_ARRAY;
            field_4_ole_classname = null;
            field_4_unknownByte = null;

        }

        /**
         * Constructs an EmbeddedObjectRef record and Sets its fields appropriately.
         *
         * @param in the record input stream.
         */
        public EmbeddedObjectRefSubRecord(ILittleEndianInput in1, int size)
        {
            // Much guess-work going on here due to lack of any documentation.
            // See similar source code in OOO:
            // http://lxr.go-oo.org/source/sc/sc/source/filter/excel/xiescher.cxx
            // 1223 void XclImpOleObj::ReadPictFmla( XclImpStream& rStrm, sal_uInt16 nRecSize )

            int streamIdOffset = in1.ReadShort(); // OOO calls this 'nFmlaLen'
            int remaining = size - LittleEndianConsts.SHORT_SIZE;

            int dataLenAfterFormula = remaining - streamIdOffset;
            int formulaSize = in1.ReadUShort();

            remaining -= LittleEndianConsts.SHORT_SIZE;
            field_1_unknown_int = in1.ReadInt();
            remaining -= LittleEndianConsts.INT_SIZE;
            byte[] formulaRawBytes = ReadRawData(in1, formulaSize);
            remaining -= formulaSize;
            field_2_refPtg = ReadRefPtg(formulaRawBytes);
            if (field_2_refPtg == null)
            {
                // common case
                // field_2_n16 seems to be 5 here
                // The formula almost looks like tTbl but the row/column values seem like garbage.
                field_2_unknownFormulaData = formulaRawBytes;
            }
            else
            {
                field_2_unknownFormulaData = null;
            }


            int stringByteCount;
            if (remaining >= dataLenAfterFormula + 3)
            {
                int tag = in1.ReadByte();
                stringByteCount = LittleEndianConsts.BYTE_SIZE;
                if (tag != 0x03)
                {
                    throw new RecordFormatException("Expected byte 0x03 here");
                }
                int nChars = in1.ReadUShort();
                stringByteCount += LittleEndianConsts.SHORT_SIZE;
                if (nChars > 0)
                {
                    // OOO: the 4th way Xcl stores a unicode string: not even a Grbit byte present if Length 0
                    field_3_unicode_flag = (in1.ReadByte() & 0x01) != 0;
                    stringByteCount += LittleEndianConsts.BYTE_SIZE;
                    if (field_3_unicode_flag)
                    {
                        field_4_ole_classname = StringUtil.ReadUnicodeLE(in1,nChars);
                        stringByteCount += nChars * 2;
                    }
                    else
                    {
                        field_4_ole_classname = StringUtil.ReadCompressedUnicode(in1,nChars);
                        stringByteCount += nChars;
                    }
                }
                else
                {
                    field_4_ole_classname = "";
                }
            }
            else
            {
                field_4_ole_classname = null;
                stringByteCount = 0;
            }
            remaining -= stringByteCount;
            // Pad to next 2-byte boundary
            if (((stringByteCount + formulaSize) % 2) != 0)
            {
                int b = in1.ReadByte();
                remaining -= LittleEndianConsts.BYTE_SIZE;
                if (field_2_refPtg != null && field_4_ole_classname == null)
                {
                    field_4_unknownByte = (byte)b;
                }
            }
            int nUnexpectedPadding = remaining - dataLenAfterFormula;

            if (nUnexpectedPadding > 0)
            {
                logger.Log(POILogger.ERROR, "Discarding " + nUnexpectedPadding + " unexpected padding bytes ");
                ReadRawData(in1, nUnexpectedPadding);
                remaining -= nUnexpectedPadding;
            }

            // Fetch the stream ID
            if (dataLenAfterFormula >= 4)
            {
                field_5_stream_id = in1.ReadInt();
                remaining -= LittleEndianConsts.INT_SIZE;
            }
            else
            {
                field_5_stream_id = null;
            }

            field_6_unknown = ReadRawData(in1, remaining);
        }



        public override short Sid
        {
            get { return sid; }
        }

        private static Ptg ReadRefPtg(byte[] formulaRawBytes)
        {
            using (MemoryStream ms = new MemoryStream(formulaRawBytes))
            {
                ILittleEndianInput in1 = new LittleEndianInputStream(ms);
                byte ptgSid = (byte)in1.ReadByte();
                switch (ptgSid)
                {
                    case AreaPtg.sid: return new AreaPtg(in1);
                    case Area3DPtg.sid: return new Area3DPtg(in1);
                    case RefPtg.sid: return new RefPtg(in1);
                    case Ref3DPtg.sid: return new Ref3DPtg(in1);
                }
                return null;
            }
        }

        private static byte[] ReadRawData(ILittleEndianInput in1, int size)
        {
            if (size < 0)
            {
                throw new ArgumentException("Negative size (" + size + ")");
            }
            if (size == 0)
            {
                return EMPTY_BYTE_ARRAY;
            }
            byte[] result = new byte[size];
            in1.ReadFully(result);
            return result;
        }

        private int GetStreamIDOffset(int formulaSize)
        {
            int result = 2 + 4; // formulaSize + f2unknown_int
            result += formulaSize;

            int stringLen;
            if (field_4_ole_classname == null)
            {
                // don't write 0x03, stringLen, flag, text
                stringLen = 0;
            }
            else
            {
                result += 1 + 2;  // 0x03, stringLen, flag
                stringLen = field_4_ole_classname.Length;
                if (stringLen > 0)
                {
                    result += 1; // flag
                    if (field_3_unicode_flag)
                    {
                        result += stringLen * 2;
                    }
                    else
                    {
                        result += stringLen;
                    }
                }
            }
            // pad to next 2 byte boundary
            if ((result % 2) != 0)
            {
                result++;
            }
            return result;
        }

        private int GetDataSize(int idOffset)
        {

            int result = 2 + idOffset; // 2 for idOffset short field itself
            if (field_5_stream_id != null)
            {
                result += 4;
            }
            return result + field_6_unknown.Length;
        }
        public override int DataSize
        {
            get
            {
                int formulaSize = field_2_refPtg == null ? field_2_unknownFormulaData.Length : field_2_refPtg.Size;
                int idOffset = GetStreamIDOffset(formulaSize);
                return GetDataSize(idOffset);
            }
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            int formulaSize = field_2_refPtg == null ? field_2_unknownFormulaData.Length : field_2_refPtg.Size;
            int idOffset = GetStreamIDOffset(formulaSize);
            int dataSize = GetDataSize(idOffset);


            out1.WriteShort(sid);
            out1.WriteShort(dataSize);

            out1.WriteShort(idOffset);
            out1.WriteShort(formulaSize);
            out1.WriteInt(field_1_unknown_int);

            int pos = 12;

            if (field_2_refPtg == null)
            {
                out1.Write(field_2_unknownFormulaData);
            }
            else
            {
                field_2_refPtg.Write(out1);
            }
            pos += formulaSize;

            int stringLen;
            if (field_4_ole_classname == null)
            {
                // don't write 0x03, stringLen, flag, text
                stringLen = 0;
            }
            else
            {
                out1.WriteByte(0x03);
                pos += 1;
                stringLen = field_4_ole_classname.Length;
                out1.WriteShort(stringLen);
                pos += 2;
                if (stringLen > 0)
                {
                    out1.WriteByte(field_3_unicode_flag ? 0x01 : 0x00);
                    pos += 1;

                    if (field_3_unicode_flag)
                    {
                        StringUtil.PutUnicodeLE(field_4_ole_classname, out1);
                        pos += stringLen * 2;
                    }
                    else
                    {
                        StringUtil.PutCompressedUnicode(field_4_ole_classname, out1);
                        pos += stringLen;
                    }
                }
            }

            // pad to next 2-byte boundary (requires 0 or 1 bytes)
            switch (idOffset - (pos - 6 ))
            { // 6 for 3 shorts: sid, dataSize, idOffset
                case 1:
                    out1.WriteByte(field_4_unknownByte == null ? 0x00 : (int)Convert.ToByte(field_4_unknownByte, CultureInfo.InvariantCulture));
                    pos++;
                    break;
                case 0:
                    break;
                default:
                    throw new InvalidOperationException("Bad padding calculation (" + idOffset + ", " + pos + ")");
            }

            if (field_5_stream_id != null)
            {
                out1.WriteInt(Convert.ToInt32(field_5_stream_id, CultureInfo.InvariantCulture));
                pos += 4;
            }
            out1.Write(field_6_unknown);
        }


        /**
         * Gets the stream ID containing the actual data.  The data itself
         * can be found under a top-level directory entry in the OLE2 filesystem
         * under the name "MBD<var>xxxxxxxx</var>" where <var>xxxxxxxx</var> is
         * this ID converted into hex (in big endian order, funnily enough.)
         *
         * @return the data stream ID. Possibly <c>null</c>
         */
        public int? StreamId
        {
            get
            {
                return field_5_stream_id;
            }
        }

        public String OLEClassName
        {
            get
            {
                return field_4_ole_classname;
            }
            set { field_4_ole_classname = value; }
        }

        public byte[] ObjectData
        {
            get
            {
                return field_6_unknown;
            }
            set
            {
                field_6_unknown = value;
            }
        }


        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("[ftPictFmla]\n");
            sb.Append("    .f2unknown     = ").Append(HexDump.IntToHex(field_1_unknown_int)).Append("\n");
            if (field_2_refPtg == null)
            {
                sb.Append("    .f3unknown     = ").Append(HexDump.ToHex(field_2_unknownFormulaData)).Append("\n");
            }
            else
            {
                sb.Append("    .formula       = ").Append(field_2_refPtg.ToString()).Append("\n");
            }
            if (field_4_ole_classname != null)
            {
                sb.Append("    .unicodeFlag   = ").Append(field_3_unicode_flag).Append("\n");
                sb.Append("    .oleClassname  = ").Append(field_4_ole_classname).Append("\n");
            }
            if (field_4_unknownByte != null)
            {
                sb.Append("    .f4unknown   = ").Append(HexDump.ByteToHex(Convert.ToByte(field_4_unknownByte, CultureInfo.InvariantCulture))).Append("\n");
            }
            if (field_5_stream_id != null)
            {
                sb.Append("    .streamId      = ").Append(HexDump.IntToHex(Convert.ToInt32(field_5_stream_id, CultureInfo.InvariantCulture))).Append("\n");
            }
            if (field_6_unknown.Length > 0)
            {
                sb.Append("    .f7unknown     = ").Append(HexDump.ToHex(field_6_unknown)).Append("\n");
            }
            sb.Append("[/ftPictFmla]");
            return sb.ToString();
        }
        public override Object Clone()
        {
            return this; // TODO proper clone
        }
        public void SetUnknownFormulaData(byte[] formularData)
        {
            field_2_unknownFormulaData = formularData;
        }

        public void SetStorageId(int storageId)
        {
            field_5_stream_id = storageId;
        }
    }
}