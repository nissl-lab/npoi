
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
    using System.Text;
    using NPOI.Util;
    using System.IO;

    /// <summary>
    /// @author Glen Stampoultzis
    /// @version $Id: EscherBlipRecord.java 569827 2007-08-26 15:26:29Z yegor $
    /// </summary>
    public class EscherBlipRecord : EscherRecord
    {
        public const short RECORD_ID_START = unchecked((short)0xF018);
        public const short RECORD_ID_END = unchecked((short)0xF117);
        public const String RECORD_DESCRIPTION = "msofbtBlip";

        private const int HEADER_SIZE = 8;

        protected byte[] field_pictureData;

        public EscherBlipRecord()
        {
        }

        /// <summary>
        /// This method deSerializes the record from a byte array.
        /// </summary>
        /// <param name="data">The byte array containing the escher record information</param>
        /// <param name="offset">The starting offset into </param>
        /// <param name="recordFactory">May be null since this is not a container record.</param>
        /// <returns>The number of bytes Read from the byte array.</returns>
        public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory)
        {
            int bytesAfterHeader = ReadHeader(data, offset);
            int pos = offset + HEADER_SIZE;

            field_pictureData = new byte[bytesAfterHeader];
            Array.Copy(data, pos, field_pictureData, 0, bytesAfterHeader);

            return bytesAfterHeader + 8;
        }

        /// <summary>
        /// Serializes the record to an existing byte array.
        /// </summary>
        /// <param name="offset"> the offset within the byte array</param>
        /// <param name="data">the data array to Serialize to</param>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        /// <returns>the number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);

            Array.Copy(field_pictureData, 0, data, offset + 4, field_pictureData.Length);

            listener.AfterRecordSerialize(offset + 4 + field_pictureData.Length, RecordId, field_pictureData.Length + 4, this);
            return field_pictureData.Length + 4;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return field_pictureData.Length + HEADER_SIZE; }
        }

        /// <summary>
        /// The short name for this record
        /// </summary>
        /// <value></value>
        public override String RecordName
        {
            get { return "Blip"; }
        }

        /// <summary>
        /// Gets or sets the picture data.
        /// </summary>
        /// <value>The picture data.</value>
        public byte[] PictureData
        {
            get { return field_pictureData; }
            set { field_pictureData = value; }
        }


        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            String nl = Environment.NewLine;

            String extraData = string.Empty;
            using (MemoryStream b = new MemoryStream())
            {
                try
                {
                    HexDump.Dump(this.field_pictureData, 0, b, 0);
                    extraData = HexDump.ToHex(b.ToArray());
                }
                catch (Exception e)
                {
                    extraData = e.ToString();
                }
                return this.GetType().Name + ":" + nl +
                        "  RecordId: 0x" + HexDump.ToHex(RecordId) + nl +
                        "  Options: 0x" + HexDump.ToHex(Options) + nl +
                        "  Version: 0x" + HexDump.ToHex(Version) + nl +
                        "  Instance: 0x" + HexDump.ToHex(Instance) + nl +
                        "  Extra Data:" + nl + extraData;
            }
        }

        public override String ToXml(String tab)
        {
            String extraData = HexDump.ToHex(field_pictureData, 32);
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                    .Append(tab).Append("\t").Append("<ExtraData>").Append(extraData).Append("</ExtraData>\n");
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }
    }
}
