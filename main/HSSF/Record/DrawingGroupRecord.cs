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
    using System.Collections;
    using NPOI.Util;
    using NPOI.DDF;

    public class DrawingGroupRecord : AbstractEscherHolderRecord
    {
        public const short sid = 0xEB;

        private const int MAX_RECORD_SIZE = 8228;
        private const int MAX_DATA_SIZE = MAX_RECORD_SIZE - 4;

        public DrawingGroupRecord()
        {
        }

        public DrawingGroupRecord(RecordInputStream in1)
            : base(in1)
        {

        }

        protected override String RecordName
        {
            get { return "MSODRAWINGGROUP"; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override int Serialize(int offset, byte [] data)
        {
            byte[] rawData = RawData;
            if (EscherRecords.Count == 0 && rawData != null)
            {
                return WriteData(offset, data, rawData);
            }
            else
            {
                byte[] buffer = new byte[RawDataSize];
                int pos = 0;
                for (IEnumerator iterator = EscherRecords.GetEnumerator(); iterator.MoveNext(); )
                {
                    EscherRecord r = (EscherRecord)iterator.Current;
                    pos += r.Serialize(pos, buffer, new NullEscherSerializationListener());
                }

                return WriteData(offset, data, buffer);
            }
        }

        /**
         * Process the bytes into escher records.
         * (Not done by default in case we break things,
         *  Unless you Set the "poi.deSerialize.escher" 
         *  system property)
         */
        public void ProcessChildRecords()
        {
            ConvertRawBytesToEscherRecords();
        }

        /**
         * Size of record (including 4 byte headers for all sections)
         */
        public override int RecordSize
        {
            get { return GrossSizeFromDataSize(this.RawDataSize); }
        }

        public int RawDataSize
        {
            get
            {
                IList escherRecords = EscherRecords;
                byte[] rawData = RawData;
                if (escherRecords.Count == 0 && rawData != null)
                {
                    return rawData.Length;
                }
                else
                {
                    int size = 0;
                    for (IEnumerator iterator = escherRecords.GetEnumerator(); iterator.MoveNext(); )
                    {
                        EscherRecord r = (EscherRecord)iterator.Current;
                        size += r.RecordSize;
                    }
                    return size;
                }
            }
        }

        public static int GrossSizeFromDataSize(int dataSize)
        {
            return dataSize + ((dataSize - 1) / MAX_DATA_SIZE + 1) * 4;
        }

        private int WriteData(int offset, byte[] data, byte[] rawData)
        {
            int writtenActualData = 0;
            int writtenRawData = 0;
            while (writtenRawData < rawData.Length)
            {
                int segmentLength = Math.Min(rawData.Length - writtenRawData, MAX_DATA_SIZE);
                if (writtenRawData / MAX_DATA_SIZE >= 2)
                    WriteContinueHeader(data, offset, segmentLength);
                else
                    WriteHeader(data, offset, segmentLength);
                writtenActualData += 4;
                offset += 4;
                Array.Copy(rawData, writtenRawData, data, offset, segmentLength);
                offset += segmentLength;
                writtenRawData += segmentLength;
                writtenActualData += segmentLength;
            }
            return writtenActualData;
        }

        private void WriteHeader(byte[] data, int offset, int sizeExcludingHeader)
        {
            LittleEndian.PutShort(data, 0 + offset, Sid);
            LittleEndian.PutShort(data, 2 + offset, (short)sizeExcludingHeader);
        }

        private void WriteContinueHeader(byte[] data, int offset, int sizeExcludingHeader)
        {
            LittleEndian.PutShort(data, 0 + offset, ContinueRecord.sid);
            LittleEndian.PutShort(data, 2 + offset, (short)sizeExcludingHeader);
        }

    }
}