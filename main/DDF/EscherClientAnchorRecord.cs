
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
    /// The escher client anchor specifies which rows and cells the shape is bound to as well as
    /// the offsets within those cells.  Each cell is 1024 units wide by 256 units long regardless
    /// of the actual size of the cell.  The EscherClientAnchorRecord only applies to the top-most
    /// shapes.  Shapes contained in groups are bound using the EscherChildAnchorRecords.
    /// @author Glen Stampoultzis
    /// </summary>
    public class EscherClientAnchorRecord : EscherRecord
    {
        public const short RECORD_ID = unchecked((short)0xF010);
        public const String RECORD_DESCRIPTION = "MsofbtClientAnchor";
        /**
         * bit[0] -  fMove (1 bit): A bit that specifies whether the shape will be kept intact when the cells are moved.
         * bit[1] - fSize (1 bit): A bit that specifies whether the shape will be kept intact when the cells are resized. If fMove is 1, the value MUST be 1.
         * bit[2-4] - reserved, MUST be 0 and MUST be ignored
         * bit[5-15]- Undefined and MUST be ignored.
         *
         * it can take values: 0, 2, 3
         */

        private short field_1_flag;
        private short field_2_col1;
        private short field_3_dx1;
        private short field_4_row1;
        private short field_5_dy1;
        private short field_6_col2;
        private short field_7_dx2;
        private short field_8_row2;
        private short field_9_dy2;
        private byte[] remainingData;
        private bool shortRecord = false;

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

            // Always find 4 two byte entries. Sometimes find 9
            if (bytesRemaining == 4) // Word format only 4 bytes
            {
                // Not sure exactly what the format is quite yet, likely a reference to a PLC
            }
            else
            {
                field_1_flag = LittleEndian.GetShort(data, pos + size); size += 2;
                field_2_col1 = LittleEndian.GetShort(data, pos + size); size += 2;
                field_3_dx1 = LittleEndian.GetShort(data, pos + size); size += 2;
                field_4_row1 = LittleEndian.GetShort(data, pos + size); size += 2;
                if (bytesRemaining >= 18)
                {
                    field_5_dy1 = LittleEndian.GetShort(data, pos + size); size += 2;
                    field_6_col2 = LittleEndian.GetShort(data, pos + size); size += 2;
                    field_7_dx2 = LittleEndian.GetShort(data, pos + size); size += 2;
                    field_8_row2 = LittleEndian.GetShort(data, pos + size); size += 2;
                    field_9_dy2 = LittleEndian.GetShort(data, pos + size); size += 2;
                    shortRecord = false;
                }
                else
                {
                    shortRecord = true;
                }
            }
            bytesRemaining -= size;
            remainingData = new byte[bytesRemaining];
            Array.Copy(data, pos + size, remainingData, 0, bytesRemaining);
            return 8 + size + bytesRemaining;
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

            if (remainingData == null) remainingData = new byte[0];
            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            int remainingBytes = remainingData.Length + (shortRecord ? 8 : 18);
            LittleEndian.PutInt(data, offset + 4, remainingBytes);
            LittleEndian.PutShort(data, offset + 8, field_1_flag);
            LittleEndian.PutShort(data, offset + 10, field_2_col1);
            LittleEndian.PutShort(data, offset + 12, field_3_dx1);
            LittleEndian.PutShort(data, offset + 14, field_4_row1);
            if (!shortRecord)
            {
                LittleEndian.PutShort(data, offset + 16, field_5_dy1);
                LittleEndian.PutShort(data, offset + 18, field_6_col2);
                LittleEndian.PutShort(data, offset + 20, field_7_dx2);
                LittleEndian.PutShort(data, offset + 22, field_8_row2);
                LittleEndian.PutShort(data, offset + 24, field_9_dy2);
            }
            Array.Copy(remainingData, 0, data, offset + (shortRecord ? 16 : 26), remainingData.Length);
            int pos = offset + 8 + (shortRecord ? 8 : 18) + remainingData.Length;

            listener.AfterRecordSerialize(pos, RecordId, pos - offset, this);
            return pos - offset;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + (shortRecord ? 8 : 18) + (remainingData == null ? 0 : remainingData.Length); }
        }

        /// <summary>
        /// The record id for this record.
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
            get { return "ClientAnchor"; }
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
                       "  Flag: " + field_1_flag + nl +
                       "  Col1: " + field_2_col1 + nl +
                       "  DX1: " + field_3_dx1 + nl +
                       "  Row1: " + field_4_row1 + nl +
                       "  DY1: " + field_5_dy1 + nl +
                       "  Col2: " + field_6_col2 + nl +
                       "  DX2: " + field_7_dx2 + nl +
                       "  Row2: " + field_8_row2 + nl +
                       "  DY2: " + field_9_dy2 + nl +
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
                    extraData = "error\n";
                }
                if (extraData.Contains("No Data"))
                {
                    extraData = "No Data";
                }
                StringBuilder builder = new StringBuilder();
                builder.Append(tab)
                       .Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId),
                                                     HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                       .Append(tab).Append("\t").Append("<Flag>").Append(field_1_flag).Append("</Flag>\n")
                       .Append(tab).Append("\t").Append("<Col1>").Append(field_2_col1).Append("</Col1>\n")
                       .Append(tab).Append("\t").Append("<DX1>").Append(field_3_dx1).Append("</DX1>\n")
                       .Append(tab).Append("\t").Append("<Row1>").Append(field_4_row1).Append("</Row1>\n")
                       .Append(tab).Append("\t").Append("<DY1>").Append(field_5_dy1).Append("</DY1>\n")
                       .Append(tab).Append("\t").Append("<Col2>").Append(field_6_col2).Append("</Col2>\n")
                       .Append(tab).Append("\t").Append("<DX2>").Append(field_7_dx2).Append("</DX2>\n")
                       .Append(tab).Append("\t").Append("<Row2>").Append(field_8_row2).Append("</Row2>\n")
                       .Append(tab).Append("\t").Append("<DY2>").Append(field_9_dy2).Append("</DY2>\n")
                       .Append(tab).Append("\t").Append("<ExtraData>").Append(extraData).Append("</ExtraData>\n");
                builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
                return builder.ToString();
            }
        }
        /// <summary>
        /// Gets or sets the flag.
        /// </summary>
        /// <value>0 = Move and size with Cells, 2 = Move but don't size with cells, 3 = Don't move or size with cells.</value>
        public short Flag
        {
            get { return field_1_flag; }
            set { field_1_flag = value; }
        }

        /// <summary>
        /// Gets or sets The column number for the top-left position.  0 based.
        /// </summary>
        /// <value>The col1.</value>
        public short Col1
        {
            get { return field_2_col1; }
            set { field_2_col1 = value; }
        }

        /// <summary>
        /// Gets or sets The x offset within the top-left cell.  Range is from 0 to 1023.
        /// </summary>
        /// <value>The DX1.</value>
        public short Dx1
        {
            get { return field_3_dx1; }
            set { field_3_dx1 = value; }
        }

        /// <summary>
        /// Gets or sets The row number for the top-left corner of the shape.
        /// </summary>
        /// <value>The row1.</value>
        public short Row1
        {
            get { return field_4_row1; }
            set { this.field_4_row1 = value; }
        }


        /// <summary>
        /// Gets or sets The y offset within the top-left corner of the current shape.
        /// </summary>
        /// <value>The dy1.</value>
        public short Dy1
        {
            get { return field_5_dy1; }
            set
            {
                shortRecord = false;
                this.field_5_dy1 = value;
            }
        }


        /// <summary>
        /// Gets or sets The column of the bottom right corner of this shape.
        /// </summary>
        /// <value>The col2.</value>
        public short Col2
        {
            get { return field_6_col2; }
            set
            {
                shortRecord = false;
                this.field_6_col2 = value;
            }
        }


        /// <summary>
        /// Gets or sets The x offset withing the cell for the bottom-right corner of this shape.
        /// </summary>
        /// <value>The DX2.</value>
        public short Dx2
        {
            get { return field_7_dx2; }
            set
            {
                shortRecord = false;
                this.field_7_dx2 = value;
            }
        }

        /// <summary>
        /// Gets or sets The row number for the bottom-right corner of the current shape.
        /// </summary>
        /// <value>The row2.</value>
        public short Row2
        {
            get { return field_8_row2; }
            set
            {
                shortRecord = false;
                this.field_8_row2 = value;
            }
        }


        /// <summary>
        /// Gets or sets The y offset withing the cell for the bottom-right corner of this shape.
        /// </summary>
        /// <value>The dy2.</value>
        public short Dy2
        {
            get { return field_9_dy2; }
            set { field_9_dy2 = value; }
        }

        /// <summary>
        /// Gets or sets the remaining data.
        /// </summary>
        /// <value>The remaining data.</value>
        public byte[] RemainingData
        {
            get { return remainingData; }
            set { remainingData = value; }
        }

    }
}