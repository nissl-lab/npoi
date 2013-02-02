/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/

using System.Text;

namespace NPOI.DDF
{
    using System;
    using System.IO;
    using NPOI.Util;

    /// <summary>
    /// @author Glen Stampoultzis
    /// @version $Id: EscherBitmapBlip.java 569827 2007-08-26 15:26:29Z yegor $
    /// </summary>
    public class EscherBitmapBlip : EscherBlipRecord
    {
        public const short RECORD_ID_JPEG = unchecked((short)0xF018) + 5;
        public const short RECORD_ID_PNG = unchecked((short)0xF018) + 6;
        public const short RECORD_ID_DIB = unchecked((short)0xF018) + 7;

        private const int HEADER_SIZE = 8;

        private byte[] field_1_UID;
        private byte field_2_marker = (byte)0xFF;


        /// <summary>
        /// This method deSerializes the record from a byte array.    
        /// </summary>
        /// <param name="data"> The byte array containing the escher record information</param>
        /// <param name="offset">The starting offset into </param>
        /// <param name="recordFactory">May be null since this is not a container record.</param>
        /// <returns>The number of bytes Read from the byte array.</returns>
        public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory)
        {
            int bytesAfterHeader = ReadHeader(data, offset);
            int pos = offset + HEADER_SIZE;

            field_1_UID = new byte[16];
            Array.Copy(data, pos, field_1_UID, 0, 16); pos += 16;
            field_2_marker = data[pos]; pos++;

            field_pictureData = new byte[bytesAfterHeader - 17];
            Array.Copy(data, pos, field_pictureData, 0, field_pictureData.Length);

            return bytesAfterHeader + HEADER_SIZE;
        }

        /// <summary>
        /// Serializes the record to an existing byte array.
        /// </summary>
        /// <param name="offset">the offset within the byte array</param>
        /// <param name="data">the data array to Serialize to</param>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        /// <returns>the number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            LittleEndian.PutInt(data, offset + 4, RecordSize - HEADER_SIZE);
            int pos = offset + HEADER_SIZE;

            Array.Copy(field_1_UID, 0, data, pos, 16);
            data[pos + 16] = field_2_marker;
            Array.Copy(field_pictureData, 0, data, pos + 17, field_pictureData.Length);

            listener.AfterRecordSerialize(offset + RecordSize, RecordId, RecordSize, this);
            return HEADER_SIZE + 16 + 1 + field_pictureData.Length;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value> Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + 16 + 1 + field_pictureData.Length; }
        }

        /// <summary>
        /// Gets or sets the UID.
        /// </summary>
        /// <value>The UID.</value>
        public byte[] UID
        {
            get { return field_1_UID; }
            set { this.field_1_UID = value; }
        }

        /// <summary>
        /// Gets or sets the marker.
        /// </summary>
        /// <value>The marker.</value>
        public byte Marker
        {
            get { return field_2_marker; }
            set { this.field_2_marker = value; }
        }

        /// <summary>
        /// Toes the string.
        /// </summary>
        /// <returns></returns>
        public override String ToString()
        {
            String nl = Environment.NewLine;

            String extraData;
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
                        "  Version: 0x" + HexDump.ToHex(Version) + nl +
                        "  Instance: 0x" + HexDump.ToHex(Instance) + nl +
                        "  UID: 0x" + HexDump.ToHex(field_1_UID) + nl +
                        "  Marker: 0x" + HexDump.ToHex(field_2_marker) + nl +
                        "  Extra Data:" + nl + extraData;
            }
        }

        public override String ToXml(String tab)
        {
            String extraData;
            //MemoryStream b = new MemoryStream();
            try
            {
                //HexDump.Dump(this.field_pictureData, 0, b, 0);
                extraData = HexDump.ToHex(this.field_pictureData);
            }
            catch (Exception e)
            {
                extraData = e.ToString();
            }
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                    .Append(tab).Append("\t").Append("<UID>0x").Append(HexDump.ToHex(field_1_UID)).Append("</UID>\n")
                    .Append(tab).Append("\t").Append("<Marker>0x").Append(HexDump.ToHex(field_2_marker)).Append("</Marker>\n")
                    .Append(tab).Append("\t").Append("<ExtraData>").Append(extraData).Append("</ExtraData>\n");
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }

    }
}