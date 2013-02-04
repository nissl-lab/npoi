
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
    /// ToGether the the EscherOptRecord this record defines some of the basic
    /// properties of a shape.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class EscherSpRecord : EscherRecord
    {
        public const short RECORD_ID = unchecked((short)0xF00A);
        public const String RECORD_DESCRIPTION = "MsofbtSp";

        public const int FLAG_GROUP = 0x0001;
        public const int FLAG_CHILD = 0x0002;
        public const int FLAG_PATRIARCH = 0x0004;
        public const int FLAG_DELETED = 0x0008;
        public const int FLAG_OLESHAPE = 0x0010;
        public const int FLAG_HAVEMASTER = 0x0020;
        public const int FLAG_FLIPHORIZ = 0x0040;
        public const int FLAG_FLIPVERT = 0x0080;
        public const int FLAG_CONNECTOR = 0x0100;
        public const int FLAG_HAVEANCHOR = 0x0200;
        public const int FLAG_BACKGROUND = 0x0400;
        public const int FLAG_HASSHAPETYPE = 0x0800;

        private int field_1_shapeId;
        private int field_2_flags;

        /// <summary>
        /// The contract of this method is to deSerialize an escher record including
        /// it's children.
        /// </summary>
        /// <param name="data">The byte array containing the Serialized escher
        /// records.</param>
        /// <param name="offset">The offset into the byte array.</param>
        /// <param name="recordFactory">A factory for creating new escher records</param>
        /// <returns>The number of bytes written.</returns>  
        public override int FillFields(byte[] data, int offset, IEscherRecordFactory recordFactory)
        {
            int bytesRemaining = ReadHeader(data, offset);
            int pos = offset + 8;
            int size = 0;
            field_1_shapeId = LittleEndian.GetInt(data, pos + size); size += 4;
            field_2_flags = LittleEndian.GetInt(data, pos + size); size += 4;
            //        bytesRemaining -= size;
            //        remainingData  =  new byte[bytesRemaining];
            //        Array.Copy( data, pos + size, remainingData, 0, bytesRemaining );
            return RecordSize;
        }

        /// <summary>
        /// Serializes to an existing byte array without serialization listener.
        /// This is done by delegating to Serialize(int, byte[], EscherSerializationListener).
        /// </summary>
        /// <param name="offset">the offset within the data byte array.</param>
        /// <param name="data"> the data array to Serialize to.</param>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        /// <returns>The number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);
            
            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            int remainingBytes = 8;
            LittleEndian.PutInt(data, offset + 4, remainingBytes);
            LittleEndian.PutInt(data, offset + 8, field_1_shapeId);
            LittleEndian.PutInt(data, offset + 12, field_2_flags);

            listener.AfterRecordSerialize(offset + RecordSize, RecordId, RecordSize, this);
            return 8 + 8;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override  int RecordSize
        {
            get{return 8 + 8;}
        }

        /// <summary>
        /// @return  the 16 bit identifier for this record.
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
            get { return "Sp"; }
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

            return this.GetType().Name + ":" + nl +
                    "  RecordId: 0x" + HexDump.ToHex(RECORD_ID) + nl +
                    "  Version: 0x" + HexDump.ToHex(Version) + nl +
                    "  ShapeType: 0x" + HexDump.ToHex(ShapeType) + nl +
                    "  ShapeId: " + field_1_shapeId + nl +
                    "  Flags: " + DecodeFlags(field_2_flags) + " (0x" + HexDump.ToHex(field_2_flags) + ")" + nl;

        }
        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                    .Append(tab).Append("\t").Append("<ShapeType>0x").Append(HexDump.ToHex(ShapeType)).Append("</ShapeType>\n")
                    .Append(tab).Append("\t").Append("<ShapeId>").Append(field_1_shapeId).Append("</ShapeId>\n")
                    .Append(tab).Append("\t").Append("<Flags>").Append(DecodeFlags(field_2_flags) + " (0x" + HexDump.ToHex(field_2_flags) + ")").Append("</Flags>\n");
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }
        /// <summary>
        /// Converts the shape flags into a more descriptive name.
        /// </summary>
        /// <param name="flags">The flags.</param>
        /// <returns></returns>
        private String DecodeFlags(int flags)
        {
            StringBuilder result = new StringBuilder();
            result.Append((flags & FLAG_GROUP) != 0 ? "|GROUP" : "");
            result.Append((flags & FLAG_CHILD) != 0 ? "|CHILD" : "");
            result.Append((flags & FLAG_PATRIARCH) != 0 ? "|PATRIARCH" : "");
            result.Append((flags & FLAG_DELETED) != 0 ? "|DELETED" : "");
            result.Append((flags & FLAG_OLESHAPE) != 0 ? "|OLESHAPE" : "");
            result.Append((flags & FLAG_HAVEMASTER) != 0 ? "|HAVEMASTER" : "");
            result.Append((flags & FLAG_FLIPHORIZ) != 0 ? "|FLIPHORIZ" : "");
            result.Append((flags & FLAG_FLIPVERT) != 0 ? "|FLIPVERT" : "");
            result.Append((flags & FLAG_CONNECTOR) != 0 ? "|CONNECTOR" : "");
            result.Append((flags & FLAG_HAVEANCHOR) != 0 ? "|HAVEANCHOR" : "");
            result.Append((flags & FLAG_BACKGROUND) != 0 ? "|BACKGROUND" : "");
            result.Append((flags & FLAG_HASSHAPETYPE) != 0 ? "|HASSHAPETYPE" : "");

            //need to check, else blows up on some records - bug 34435
            if (result.Length > 0)
            {
                result.Remove(0,1);
            }
            return result.ToString();
        }

        /// <summary>
        /// Gets or sets A number that identifies this shape
        /// </summary>
        /// <value>The shape id.</value>
        public int ShapeId
        {
            get { return field_1_shapeId; }
            set { this.field_1_shapeId = value; }
        }

        /// <summary>
        /// The flags that apply to this shape.
        /// </summary>
        /// <value>The flags.</value>
        public int Flags
        {
            get { return field_2_flags; }
            set { this.field_2_flags = value; }
        }

        /// <summary>
        /// Get or set shape type. Must be one of MSOSPT values (see [MS-ODRAW] for details).
        /// </summary>
        public short ShapeType
        {
            get { return Instance; }
            set { Instance = (value); }
        }
    }
}