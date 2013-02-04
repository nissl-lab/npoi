
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.DDF
{
    using System;
    using System.IO;
    using System.Text;
    using NPOI.Util;


    /// <summary>
    /// The EscherClientDataRecord is used to store client specific data about the position of a
    /// shape within a container.
    /// @author Glen Stampoultzis
    /// </summary>
    public class EscherClientDataRecord : EscherRecord
    {
        public const short RECORD_ID = unchecked((short)0xF011);
        public const String RECORD_DESCRIPTION = "MsofbtClientData";

        private byte[] remainingData;

        /// <summary>
        /// This method deSerializes the record from a byte array.
        /// </summary>
        /// <param name="data">The byte array containing the escher record information</param>
        /// <param name="offset">The starting offset into data</param>
        /// <param name="recordFactory">May be null since this is not a container record.</param>
        /// <returns>The number of bytes Read from the byte array.</returns>
        public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory)
        {
            int bytesRemaining = ReadHeader(data, offset);
            int pos = offset + 8;
            remainingData = new byte[bytesRemaining];
            Array.Copy(data, pos, remainingData, 0, bytesRemaining);
            return 8 + bytesRemaining;
        }

        /**
         * This method Serializes this escher record into a byte array.
         *
         * @param offset   The offset into <c>data</c> to start writing the record data to.
         * @param data     The byte array to Serialize to.
         * @param listener A listener to retrieve start and end callbacks.  Use a <c>NullEscherSerailizationListener</c> to ignore these events.
         * @return The number of bytes written.
         * @see NullEscherSerializationListener
         */
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            if (remainingData == null) remainingData = new byte[0];
            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            LittleEndian.PutInt(data, offset + 4, remainingData.Length);
            Array.Copy(remainingData, 0, data, offset + 8, remainingData.Length);
            int pos = offset + 8 + remainingData.Length;

            listener.AfterRecordSerialize(pos, RecordId, pos - offset, this);
            return pos - offset;
        }

        /**
         * Returns the number of bytes that are required to Serialize this record.
         *
         * @return Number of bytes
         */
        public override int RecordSize
        {
            get { return 8 + (remainingData == null ? 0 : remainingData.Length); }
        }

        /**
         * Returns the identifier of this record.
         */
        public override short RecordId
        {
            get { return RECORD_ID; }
        }

        /**
         * The short name for this record
         */
        public override String RecordName
        {
            get { return "ClientData"; }
        }

        /**
         * Returns the string representation of this record.
         */
        public override String ToString()
        {
            String nl = Environment.NewLine;

            String extraData;
            using (MemoryStream b = new MemoryStream())
            {
                try
                {
                    HexDump.Dump(this.remainingData, 0, b, 0);
                    //extraData = b.ToString();
                    extraData = Encoding.UTF8.GetString(b.ToArray());
                }
                catch (Exception)
                {
                    extraData = "error\n";
                }
                return GetType().Name + ":" + nl +
                        "  RecordId: 0x" + HexDump.ToHex(RECORD_ID) + nl +
                        "  Version: 0x" + HexDump.ToHex(Version) + nl +
                        "  Instance: 0x" + HexDump.ToHex(Instance) + nl +
                        "  Extra Data:" + nl + extraData;
            }
        }
        public override String ToXml(String tab)
        {
            String extraData;
            using (MemoryStream b = new MemoryStream())
            {
                try
                {
                    HexDump.Dump(this.remainingData, 0, b, 0);
                    extraData = HexDump.ToHex(b.ToArray());
                }
                catch (Exception)
                {
                    extraData = "error";
                }
                if (extraData.Contains("No Data"))
                {
                    extraData = "No Data";
                }
                StringBuilder builder = new StringBuilder();
                builder.Append(tab)
                       .Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId),
                                                     HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                       .Append(tab).Append("\t").Append("<ExtraData>").Append(extraData).Append("</ExtraData>\n");
                builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
                return builder.ToString();
            }
        }

        /**
         * Any data recording this record.
         */
        public byte[] RemainingData
        {
            get { return remainingData; }
            set { this.remainingData = value; }
        }
    }
}