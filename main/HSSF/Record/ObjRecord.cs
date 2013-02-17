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
    using NPOI.Util;

    using System.IO;
    using System.Collections.Generic;


    /**
     * The obj record is used to hold various graphic objects and controls.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class ObjRecord : Record
    {
        private const int NORMAL_PAD_ALIGNMENT = 2;
	    private const int MAX_PAD_ALIGNMENT = 4;

        public const short sid = 0x5D;
        private List<SubRecord> subrecords;
        	/** used when POI has no idea what is going on */
	    private byte[] _uninterpretedData;
    /**
* Excel seems to tolerate padding to quad or double byte length
*/
        private bool _isPaddedToQuadByteMultiple;

        //00000000 15 00 12 00 01 00 01 00 11 60 00 00 00 00 00 0D .........`......
        //00000010 26 01 00 00 00 00 00 00 00 00                   &.........


        public ObjRecord()
        {
            subrecords = new List<SubRecord>(2);
            // TODO - ensure 2 sub-records (ftCmo 15h, and ftEnd 00h) are always created
            _uninterpretedData = null;
        }

        /**
         * Constructs a OBJ record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public ObjRecord(RecordInputStream in1)
        {
            // TODO - problems with OBJ sub-records stream
            // MS spec says first sub-record is always CommonObjectDataSubRecord,
            // and last is
            // always EndSubRecord. OOO spec does not mention ObjRecord(0x005D).
            // Existing POI test data seems to violate that rule. Some test data
            // seems to contain
            // garbage, and a crash is only averted by stopping at what looks like
            // the 'EndSubRecord'

            //Check if this can be continued, if so then the
            //following wont work properly
            //int subSize = 0;
            byte[] subRecordData = in1.ReadRemainder();

            if (LittleEndian.GetUShort(subRecordData, 0) != CommonObjectDataSubRecord.sid)
            {
                // seems to occur in just one junit on "OddStyleRecord.xls" (file created by CrystalReports)
                // Excel tolerates the funny ObjRecord, and replaces it with a corrected version
                // The exact logic/reasoning is not yet understood
                _uninterpretedData = subRecordData;
                subrecords = null;
                return;
            }
            //if (subRecordData.Length % 2 != 0)
            //{
            //    String msg = "Unexpected length of subRecordData : " + HexDump.ToHex(subRecordData);
            //    throw new RecordFormatException(msg);
            //}
            subrecords = new List<SubRecord>();
            using (MemoryStream bais = new MemoryStream(subRecordData))
            {
                LittleEndianInputStream subRecStream = new LittleEndianInputStream(bais);
                CommonObjectDataSubRecord cmo = (CommonObjectDataSubRecord)SubRecord.CreateSubRecord(subRecStream, 0);
                subrecords.Add(cmo);
                while (true)
                {
                    SubRecord subRecord = SubRecord.CreateSubRecord(subRecStream, cmo.ObjectType);
                    subrecords.Add(subRecord);
                    if (subRecord.IsTerminating)
                    {
                        break;
                    }
                }
                int nRemainingBytes = subRecStream.Available();
                if (nRemainingBytes > 0)
                {
                    // At present (Oct-2008), most unit test samples have (subRecordData.length % 2 == 0)
                    _isPaddedToQuadByteMultiple = subRecordData.Length % MAX_PAD_ALIGNMENT == 0;
                    if (nRemainingBytes >= (_isPaddedToQuadByteMultiple ? MAX_PAD_ALIGNMENT : NORMAL_PAD_ALIGNMENT))
                    {
                        if (!CanPaddingBeDiscarded(subRecordData, nRemainingBytes))
                        {
                            String msg = "Leftover " + nRemainingBytes
                                + " bytes in subrecord data " + HexDump.ToHex(subRecordData);
                            throw new RecordFormatException(msg);
                        }
                        _isPaddedToQuadByteMultiple = false;
                    }
                }
                else
                {
                    _isPaddedToQuadByteMultiple = false;
                }
                _uninterpretedData = null;
            }
        }
        /**
     * Some XLS files have ObjRecords with nearly 8Kb of excessive padding. These were probably
     * written by a version of POI (around 3.1) which incorrectly interpreted the second short of
     * the ftLbs subrecord (0x1FEE) as a length, and read that many bytes as padding (other bugs
     * helped allow this to occur).
     * 
     * Excel reads files with this excessive padding OK, truncating the over-sized ObjRecord back
     * to the its proper size.  POI does the same.
     */
        private static bool CanPaddingBeDiscarded(byte[] data, int nRemainingBytes)
        {
            // make sure none of the padding looks important
            for (int i = data.Length - nRemainingBytes; i < data.Length; i++)
            {
                if (data[i] != 0x00)
                {
                    return false;
                }
            }
            return true;
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("[OBJ]\n");
            for (int i = 0; i < subrecords.Count; i++)
            {
                SubRecord record = subrecords[i];
                sb.Append("SUBRECORD: ").Append(record.ToString());
            }
            sb.Append("[/OBJ]\n");
            return sb.ToString();
        }

        public override int Serialize(int offset, byte [] data)
        {
            int recSize = RecordSize;
            int dataSize = recSize - 4;

            LittleEndianByteArrayOutputStream out1 = new LittleEndianByteArrayOutputStream(data, offset, recSize);

		    out1.WriteShort(sid);
		    out1.WriteShort(dataSize);

            if (_uninterpretedData == null)
            {
                for (int i = 0; i < subrecords.Count; i++)
                {
                    SubRecord record = subrecords[i];
                    record.Serialize(out1);
                }
                int expectedEndIx = offset + dataSize;
                // padding
                while (out1.WriteIndex < expectedEndIx)
                {
                    out1.WriteByte(0);
                }
            }
            else
            {
                out1.Write(_uninterpretedData);
            }
            return recSize;
        }

        /**
         * Size of record (excluding 4 byte header)
         */
        public override int RecordSize
        {
            get
            {
                if (_uninterpretedData != null)
                {
                    return _uninterpretedData.Length + 4;
                }
                int size = 0;
                for (int i = subrecords.Count - 1; i >= 0; i--)
                {
                    SubRecord record = subrecords[i];
                    size += record.DataSize + 4;
                }
                if (_isPaddedToQuadByteMultiple)
                {
                    while (size % MAX_PAD_ALIGNMENT != 0)
                    {
                        size++;
                    }
                }
                else
                {
                    while (size % NORMAL_PAD_ALIGNMENT != 0)
                    {
                        size++;
                    }
                }
                return size + 4;
            }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public List<SubRecord> SubRecords
        {
            get
            {
                return subrecords;
            }
        }

        public void ClearSubRecords()
        {
            subrecords.Clear();
        }

        public void AddSubRecord(int index, SubRecord element)
        {
            subrecords.Insert(index, element);
        }

        public void AddSubRecord(SubRecord o)
        {
            subrecords.Add(o);
        }

        public override Object Clone()
        {
            ObjRecord rec = new ObjRecord();

            for (int i = 0; i < subrecords.Count; i++)
            {
                SubRecord record = subrecords[i];
                rec.AddSubRecord((SubRecord)record.Clone());
            }

            return rec;
        }
    }
}