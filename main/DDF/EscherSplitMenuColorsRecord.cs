
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
    /// A list of the most recently used colours for the drawings contained in
    /// this document.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class EscherSplitMenuColorsRecord : EscherRecord
    {
        public const short RECORD_ID = unchecked((short)0xF11E);
        public const String RECORD_DESCRIPTION = "MsofbtSplitMenuColors";

        private int field_1_color1;
        private int field_2_color2;
        private int field_3_color3;
        private int field_4_color4;

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
            field_1_color1 = LittleEndian.GetInt(data, pos + size); size += 4;
            field_2_color2 = LittleEndian.GetInt(data, pos + size); size += 4;
            field_3_color3 = LittleEndian.GetInt(data, pos + size); size += 4;
            field_4_color4 = LittleEndian.GetInt(data, pos + size); size += 4;
            bytesRemaining -= size;
            if (bytesRemaining != 0)
                throw new RecordFormatException("Expecting no remaining data but got " + bytesRemaining + " byte(s).");
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
            //        int field_2_numIdClusters = field_5_fileIdClusters.Length + 1;
            listener.BeforeRecordSerialize(offset, RecordId, this);

            int pos = offset;
            LittleEndian.PutShort(data, pos, Options); pos += 2;
            LittleEndian.PutShort(data, pos, RecordId); pos += 2;
            int remainingBytes = RecordSize - 8;

            LittleEndian.PutInt(data, pos, remainingBytes); pos += 4;
            LittleEndian.PutInt(data, pos, field_1_color1); pos += 4;
            LittleEndian.PutInt(data, pos, field_2_color2); pos += 4;
            LittleEndian.PutInt(data, pos, field_3_color3); pos += 4;
            LittleEndian.PutInt(data, pos, field_4_color4); pos += 4;
            
            listener.AfterRecordSerialize(pos, RecordId, pos - offset, this);
            return RecordSize;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + 4 * 4; }
        }

        /// <summary>
        /// Return the current record id.
        /// </summary>
        /// <value>the 16 bit identifer for this record.</value>
        public override short RecordId
        {
            get { return RECORD_ID; }
        }

        /// <summary>
        /// Gets the short name for this record
        /// </summary>
        /// <value>The name of the record.</value>
        public override String RecordName
        {
            get { return "SplitMenuColors"; }
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        /// @return  a string representation of this record.
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
                    "  Color1: 0x" + HexDump.ToHex(field_1_color1) + nl +
                    "  Color2: 0x" + HexDump.ToHex(field_2_color2) + nl +
                    "  Color3: 0x" + HexDump.ToHex(field_3_color3) + nl +
                    "  Color4: 0x" + HexDump.ToHex(field_4_color4) + nl +
                    "";

        }
        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                    .Append(tab).Append("\t").Append("<Color1>0x").Append(HexDump.ToHex(field_1_color1)).Append("</Color1>\n")
                    .Append(tab).Append("\t").Append("<Color2>0x").Append(HexDump.ToHex(field_2_color2)).Append("</Color2>\n")
                    .Append(tab).Append("\t").Append("<Color3>0x").Append(HexDump.ToHex(field_3_color3)).Append("</Color3>\n")
                    .Append(tab).Append("\t").Append("<Color4>0x").Append(HexDump.ToHex(field_4_color4)).Append("</Color4>\n");
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }
        /// <summary>
        /// Gets or sets the color1.
        /// </summary>
        /// <value>The color1.</value>
        public int Color1
        {
            get { return field_1_color1; }
            set { field_1_color1 = value; }
        }

        /// <summary>
        /// Gets or sets the color2.
        /// </summary>
        /// <value>The color2.</value>
        public int Color2
        {
            get { return field_2_color2; }
            set { field_2_color2 = value; }
        }

        /// <summary>
        /// Gets or sets the color3.
        /// </summary>
        /// <value>The color3.</value>
        public int Color3
        {
            get { return field_3_color3; }
            set { field_3_color3 = value; }
        }

        /// <summary>
        /// Gets or sets the color4.
        /// </summary>
        /// <value>The color4.</value>
        public int Color4
        {
            get { return field_4_color4; }
            set { field_4_color4 = value; }
        }

    }
}