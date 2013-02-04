
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

    /// <summary>
    /// The escher child achor record is used to specify the position of a shape under an
    /// existing group.  The first level of shape records use a EscherClientAnchor record instead.
    /// @author Glen Stampoultzis
    /// </summary>
    public class EscherChildAnchorRecord : EscherRecord
    {
        public const short RECORD_ID = unchecked((short)0xF00F);
        public const String RECORD_DESCRIPTION = "MsofbtChildAnchor";

        private int field_1_dx1;
        private int field_2_dy1;
        private int field_3_dx2;
        private int field_4_dy2;

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
            int size = 0;
            field_1_dx1 = LittleEndian.GetInt(data, pos + size); size += 4;
            field_2_dy1 = LittleEndian.GetInt(data, pos + size); size += 4;
            field_3_dx2 = LittleEndian.GetInt(data, pos + size); size += 4;
            field_4_dy2 = LittleEndian.GetInt(data, pos + size); size += 4;
            return 8 + size;
        }

        /// <summary>
        /// This method Serializes this escher record into a byte array.
        /// </summary>
        /// <param name="offset">The offset into data to start writing the record data to.</param>
        /// <param name="data">The byte array to Serialize to.</param>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        /// <returns>The number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);
            int pos = offset;
            LittleEndian.PutShort(data, pos, Options); pos += 2;
            LittleEndian.PutShort(data, pos, RecordId); pos += 2;
            LittleEndian.PutInt(data, pos, RecordSize - 8); pos += 4;
            LittleEndian.PutInt(data, pos, field_1_dx1); pos += 4;
            LittleEndian.PutInt(data, pos, field_2_dy1); pos += 4;
            LittleEndian.PutInt(data, pos, field_3_dx2); pos += 4;
            LittleEndian.PutInt(data, pos, field_4_dy2); pos += 4;

            listener.AfterRecordSerialize(pos, RecordId, pos - offset, this);
            return pos - offset;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + 4 * 4; }
        }

        /// <summary>
        /// The record id for the EscherChildAnchorRecord.
        /// </summary>
        /// <value></value>
        public override short RecordId
        {
            get { return RECORD_ID; }
        }

        /// <summary>
        /// The short name for this record
        /// </summary>
        /// <value></value>
        public override String RecordName
        {
            get { return "ChildAnchor"; }
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

            return GetType().Name + ":" + nl +
                    "  RecordId: 0x" + HexDump.ToHex(RECORD_ID) + nl +
                    "  Version: 0x" + HexDump.ToHex(Version) + nl +
                    "  Instance: 0x" + HexDump.ToHex(Instance) + nl +
                    "  X1: " + field_1_dx1 + nl +
                    "  Y1: " + field_2_dy1 + nl +
                    "  X2: " + field_3_dx2 + nl +
                    "  Y2: " + field_4_dy2 + nl;

        }
        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                    .Append(tab).Append("\t").Append("<X1>").Append(field_1_dx1).Append("</X1>\n")
                    .Append(tab).Append("\t").Append("<Y1>").Append(field_2_dy1).Append("</Y1>\n")
                    .Append(tab).Append("\t").Append("<X2>").Append(field_3_dx2).Append("</X2>\n")
                    .Append(tab).Append("\t").Append("<Y2>").Append(field_4_dy2).Append("</Y2>\n");
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }

        /// <summary>
        /// Gets or sets offset within the parent coordinate space for the top left point.
        /// </summary>
        /// <value>The DX1.</value>
        public int Dx1
        {
            get { return field_1_dx1; }
            set { this.field_1_dx1 = value; }
        }

        /// <summary>
        /// Gets or sets the offset within the parent coordinate space for the top left point.
        /// </summary>
        /// <value>The dy1.</value>
        public int Dy1
        {
            get { return field_2_dy1; }
            set { field_2_dy1 = value; }
        }
        /// <summary>
        /// Gets or sets the offset within the parent coordinate space for the bottom right point.
        /// </summary>
        /// <value>The DX2.</value>
        public int Dx2
        {
            get { return field_3_dx2; }
            set { field_3_dx2 = value; }
        }

        /// <summary>
        /// Gets or sets the offset within the parent coordinate space for the bottom right point.
        /// </summary>
        /// <value>The dy2.</value>
        public int Dy2
        {
            get { return field_4_dy2; }
            set { field_4_dy2 = value; }
        }

    }
}