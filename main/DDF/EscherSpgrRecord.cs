
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

using System.Text;

namespace NPOI.DDF
{
    using System;
    using NPOI.Util;

    /// <summary>
    /// The spgr record defines information about a shape group.  Groups in escher
    /// are simply another form of shape that you can't physically see.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class EscherSpgrRecord : EscherRecord
    {
        public const short RECORD_ID = unchecked((short)0xF009);
        public const String RECORD_DESCRIPTION = "MsofbtSpgr";

        private int field_1_rectX1;
        private int field_2_rectY1;
        private int field_3_rectX2;
        private int field_4_rectY2;

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
            field_1_rectX1 = LittleEndian.GetInt(data, pos + size); size += 4;
            field_2_rectY1 = LittleEndian.GetInt(data, pos + size); size += 4;
            field_3_rectX2 = LittleEndian.GetInt(data, pos + size); size += 4;
            field_4_rectY2 = LittleEndian.GetInt(data, pos + size); size += 4;
            bytesRemaining -= size;
            if (bytesRemaining != 0) throw new RecordFormatException("Expected no remaining bytes but got " + bytesRemaining);
            //        remainingData  =  new byte[bytesRemaining];
            //        Array.Copy( data, pos + size, remainingData, 0, bytesRemaining );
            return 8 + size + bytesRemaining;
        }

        /// <summary>
        /// This method Serializes this escher record into a byte array
        /// </summary>
        /// <param name="offset">The offset into data
        ///  to start writing the record data to.</param>
        /// <param name="data">The byte array to Serialize to.</param>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        /// <returns>The number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);
            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            int remainingBytes = 16;
            LittleEndian.PutInt(data, offset + 4, remainingBytes);
            LittleEndian.PutInt(data, offset + 8, field_1_rectX1);
            LittleEndian.PutInt(data, offset + 12, field_2_rectY1);
            LittleEndian.PutInt(data, offset + 16, field_3_rectX2);
            LittleEndian.PutInt(data, offset + 20, field_4_rectY2);
            //        Array.Copy( remainingData, 0, data, offset + 26, remainingData.Length );
            //        int pos = offset + 8 + 18 + remainingData.Length;
            listener.AfterRecordSerialize(offset + RecordSize, RecordId, offset + RecordSize, this);
            return 8 + 16;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + 16; }
        }

        /// <summary>
        /// Return the current record id.
        /// </summary>
        /// <value>The 16 bit identifier of this shape group record.</value>
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
            get { return "Spgr"; }
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

            //        String extraData;
            //        MemoryStream b = new MemoryStream();
            //        try
            //        {
            //            HexDump.dump(this.remainingData, 0, b, 0);
            //            extraData = b.ToString();
            //        }
            //        catch ( Exception e )
            //        {
            //            extraData = "error";
            //        }
            return GetType().Name + ":" + nl +
                    "  RecordId: 0x" + HexDump.ToHex(RECORD_ID) + nl +
                    "  Version: 0x" + HexDump.ToHex(Version) + nl +
                    "  Instance: 0x" + HexDump.ToHex(Instance) + nl +
                    "  RectX: " + field_1_rectX1 + nl +
                    "  RectY: " + field_2_rectY1 + nl +
                    "  RectWidth: " + field_3_rectX2 + nl +
                    "  RectHeight: " + field_4_rectY2 + nl;

        }

        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                    .Append(tab).Append("\t").Append("<RectX>").Append(field_1_rectX1).Append("</RectX>\n")
                    .Append(tab).Append("\t").Append("<RectY>").Append(field_2_rectY1).Append("</RectY>\n")
                    .Append(tab).Append("\t").Append("<RectWidth>").Append(field_3_rectX2).Append("</RectWidth>\n")
                    .Append(tab).Append("\t").Append("<RectHeight>").Append(field_4_rectY2).Append("</RectHeight>\n");
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }
        /// <summary>
        /// Gets or sets the starting top-left coordinate of child records.
        /// </summary>
        /// <value>The rect x1.</value>
        public int RectX1
        {
            get { return field_1_rectX1; }
            set { field_1_rectX1 = value; }
        }

        /// <summary>
        /// Gets or sets the starting bottom-right coordinate of child records.
        /// </summary>
        /// <value>The rect x2.</value>
        public int RectX2
        {
            get { return field_3_rectX2; }
            set { field_3_rectX2 = value; }
        }

        /// <summary>
        /// Gets or sets the starting top-left coordinate of child records.
        /// </summary>
        /// <value>The rect y1.</value>
        public int RectY1
        {
            get { return field_2_rectY1; }
            set { field_2_rectY1 = value; }
        }
        /// <summary>
        /// Gets or sets the starting bottom-right coordinate of child records.
        /// </summary>
        /// <value>The rect y2.</value>
        public int RectY2
        {
            get { return field_4_rectY2; }
            set { field_4_rectY2 = value; }
        }


    }
}