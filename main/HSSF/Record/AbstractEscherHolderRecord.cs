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
    using System.Collections;
    using NPOI.DDF;
    using NPOI.Util;
    using System.Collections.Generic;
    using NPOI.HSSF.Util;

    /**
     * The escher container record is used to hold escher records.  It is abstract and
     * must be subclassed for maximum benefit.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     * @author Michael Zalewski (zalewski at optonline.net)
     */
    public abstract class AbstractEscherHolderRecord : Record
    {
        private static bool DESERIALISE;
        static AbstractEscherHolderRecord()
        {
            DESERIALISE = false;
            //try
            //{
            //    DESERIALISE = (System.Configuration.ConfigurationManager.AppSettings["poi.deserialize.escher"] != null);
            //}
            //catch (Exception)
            //{
            //    DESERIALISE = false;
            //}
        }

        private List<EscherRecord> escherRecords;
        private LazilyConcatenatedByteArray rawDataContainer = new LazilyConcatenatedByteArray();
        //private byte[] rawData;


        public AbstractEscherHolderRecord()
        {
            escherRecords = new List<EscherRecord>();
        }

        /**
         * Constructs a Bar record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */
        public AbstractEscherHolderRecord(RecordInputStream in1)
        {
            escherRecords = new List<EscherRecord>();
            if (!DESERIALISE)
            {
                rawDataContainer.Concatenate(in1.ReadRemainder());
            }
            else
            {
                byte[] data = in1.ReadAllContinuedRemainder();
                ConvertToEscherRecords(0, data.Length, data);
            }

        }

        protected void ConvertRawBytesToEscherRecords()
        {
            if (!DESERIALISE)
            {
                byte[] rawData = RawData;
                ConvertToEscherRecords(0, rawData.Length, rawData);
            }
        }
        private void ConvertToEscherRecords(int offset, int size, byte[] data)
        {
            escherRecords.Clear();
            IEscherRecordFactory recordFactory = new DefaultEscherRecordFactory();
            int pos = offset;
            while (pos < offset + size)
            {
                EscherRecord r = recordFactory.CreateRecord(data, pos);
                int bytesRead = r.FillFields(data, pos, recordFactory);
                escherRecords.Add(r);
                pos += bytesRead;
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            String nl = Environment.NewLine;
            buffer.Append('[' + RecordName + ']' + nl);
            if (escherRecords.Count == 0)
                buffer.Append("No Escher Records Decoded" + nl);
            foreach (EscherRecord r in escherRecords)
            {
                buffer.Append(r.ToString());
            }
            buffer.Append("[/" + RecordName + ']' + nl);

            return buffer.ToString();
        }

        protected abstract String RecordName { get; }

        public override int Serialize(int offset, byte[] data)
        {
            LittleEndian.PutShort(data, 0 + offset, Sid);
            LittleEndian.PutShort(data, 2 + offset, (short)(RecordSize - 4));
            byte[] rawData = RawData;
            if (escherRecords.Count == 0 && rawData != null)
            {
                LittleEndian.PutShort(data, 0 + offset, Sid);
                LittleEndian.PutShort(data, 2 + offset, (short)(RecordSize - 4));
                Array.Copy(rawData, 0, data, 4 + offset, rawData.Length);
                return rawData.Length + 4;
            }
            LittleEndian.PutShort(data, 0 + offset, Sid);
            LittleEndian.PutShort(data, 2 + offset, (short)(RecordSize - 4));

            int pos = offset + 4;
            foreach (EscherRecord r in escherRecords)
            {
                pos += r.Serialize(pos, data, new NullEscherSerializationListener());
            }
            return RecordSize;
        }



        /**
         * Size of record (including 4 byte header)
         */
        public override int RecordSize
        {
            get
            {
                byte[] rawData = RawData;
                if (escherRecords.Count == 0 && rawData != null)
                {
                    return rawData.Length + 4;
                }
                else
                {
                    int size = 4;
                    foreach (EscherRecord r in escherRecords)
                    {
                        //EscherRecord r = (EscherRecord)iterator.Current;
                        size += r.RecordSize;
                    }
                    return size;
                }
            }
        }

        public override object Clone()
        {
            return CloneViaReserialise();
        }


        /**
 * Clone the current record, via a call to serialise
 *  it, and another to Create a new record from the
 *  bytes.
 * May only be used for classes which don't have
 *  internal counts / ids in them. For those which
 *  do, a full record-aware serialise is needed, which
 *  allocates new ids / counts as needed.
 */
        //public override Record CloneViaReserialise()
        //{
        //    // Do it via a re-serialise
        //    // It's a cheat, but it works...
        //    byte[] b = this.Serialize();
        //    using (var ms = new System.IO.MemoryStream(b))
        //    {
        //        RecordInputStream rinp = new RecordInputStream(ms);
        //        rinp.NextRecord();

        //        Record[] r = RecordFactory.CreateRecord(rinp);
        //        if (r.Length != 1)
        //        {
        //            throw new InvalidOperationException("Re-serialised a record to Clone it, but got " + r.Length + " records back!");
        //        }
        //        return r[0];
        //    }
        //}

        public void AddEscherRecord(int index, EscherRecord element)
        {
            escherRecords.Insert(index, element);
        }

        public bool AddEscherRecord(EscherRecord element)
        {
            escherRecords.Add(element);
            return true;
        }

        public List<EscherRecord> EscherRecords
        {
            get { return escherRecords; }
        }

        public void ClearEscherRecords()
        {
            escherRecords.Clear();
        }

        /**
         * If we have a EscherContainerRecord as one of our
         *  children (and most top level escher holders do),
         *  then return that.
         */
        public EscherContainerRecord GetEscherContainer()
        {
            for (IEnumerator it = escherRecords.GetEnumerator(); it.MoveNext(); )
            {
                Object er = it.Current;
                if (er is EscherContainerRecord)
                {
                    return (EscherContainerRecord)er;
                }
            }
            return null;
        }

        /**
         * Descends into all our children, returning the
         *  first EscherRecord with the given id, or null
         *  if none found
         */
        public EscherRecord FindFirstWithId(short id)
        {
            return FindFirstWithId(id, EscherRecords);
        }
        private EscherRecord FindFirstWithId(short id, List<EscherRecord> records)
        {
            // Check at our level
            for (IEnumerator it = records.GetEnumerator(); it.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)it.Current;
                if (r.RecordId == id)
                {
                    return r;
                }
            }

            // Then Check our children in turn
            for (IEnumerator it = records.GetEnumerator(); it.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)it.Current;
                if (r.IsContainerRecord)
                {
                    EscherRecord found =
                        FindFirstWithId(id, r.ChildRecords);
                    if (found != null)
                    {
                        return found;
                    }
                }
            }

            // Not found in this lot
            return null;
        }


        public EscherRecord GetEscherRecord(int index)
        {
            return escherRecords[index];
        }

        /**
         * Big drawing Group records are split but it's easier to deal with them
         * as a whole Group so we need to join them toGether.
         */
        public void Join(AbstractEscherHolderRecord record)
        {
            rawDataContainer.Concatenate(record.RawData);
        }

        public void ProcessContinueRecord(byte[] record)
        {
            rawDataContainer.Concatenate(record);
        }

        public byte[] RawData
        {
            get { return rawDataContainer.ToArray(); }
            set
            {
                rawDataContainer.Clear();
                rawDataContainer.Concatenate(value);
            }
        }

        /**
         * Convert raw data to escher records.
         */
        public void Decode()
        {
            if (null == escherRecords || 0 == escherRecords.Count)
            {
                byte[] rawData = RawData;
                ConvertToEscherRecords(0, rawData.Length, rawData);
            }
        }

    }

}


