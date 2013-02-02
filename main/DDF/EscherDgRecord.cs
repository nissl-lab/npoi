
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
    /// This record simply holds the number of shapes in the drawing group and the
    /// last shape id used for this drawing group.
    /// @author Glen Stampoultzis
    /// </summary>
    public class EscherDgRecord : EscherRecord
    {
        public const short RECORD_ID = unchecked((short)0xF008);
        public const String RECORD_DESCRIPTION = "MsofbtDg";

        private int field_1_numShapes;
        private int field_2_lastMSOSPID;

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
            field_1_numShapes = LittleEndian.GetInt(data, pos + size); size += 4;
            field_2_lastMSOSPID = LittleEndian.GetInt(data, pos + size); size += 4;
            //        bytesRemaining -= size;
            //        remainingData  =  new byte[bytesRemaining];
            //        Array.Copy( data, pos + size, remainingData, 0, bytesRemaining );
            return RecordSize;
        }

        /// <summary>
        /// This method Serializes this escher record into a byte array.
        /// </summary>
        /// <param name="offset"> The offset into data to start writing the record data to.</param>
        /// <param name="data"> The byte array to Serialize to.</param>
        /// <returns>The number of bytes written.</returns>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            LittleEndian.PutInt(data, offset + 4, 8);
            LittleEndian.PutInt(data, offset + 8, field_1_numShapes);
            LittleEndian.PutInt(data, offset + 12, field_2_lastMSOSPID);
            //        Array.Copy( remainingData, 0, data, offset + 26, remainingData.Length );
            //        int pos = offset + 8 + 18 + remainingData.Length;

            listener.AfterRecordSerialize(offset + 16, RecordId, RecordSize, this);
            return RecordSize;
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + 8; }
        }

        /// <summary>
        /// Return the current record id.
        /// </summary>
        /// <value>The 16 bit record id.</value>
        public override short RecordId
        {
            get { return RECORD_ID; }
        }

        /// <summary>
        ///  The short name for this record
        /// </summary>
        /// <value></value>
        public override String RecordName
        {
            get { return "Dg"; }
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            String nl =Environment.NewLine;

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
                    "  NumShapes: " + field_1_numShapes + nl +
                    "  LastMSOSPID: " + field_2_lastMSOSPID + nl;

        }

        public override String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append(FormatXmlRecordHeader(GetType().Name, HexDump.ToHex(RecordId), HexDump.ToHex(Version), HexDump.ToHex(Instance)))
                    .Append(tab).Append("\t").Append("<NumShapes>").Append(field_1_numShapes).Append("</NumShapes>\n")
                    .Append(tab).Append("\t").Append("<LastMSOSPID>").Append(field_2_lastMSOSPID).Append("</LastMSOSPID>\n");
            builder.Append(tab).Append("</").Append(GetType().Name).Append(">\n");
            return builder.ToString();
        }
        /// <summary>
        /// Gets or sets The number of shapes in this drawing group.
        /// </summary>
        /// <value>The num shapes.</value>
        public int NumShapes
        {
            get { return field_1_numShapes; }
            set { field_1_numShapes = value; }
        }


        /// <summary>
        /// Gets or sets The last shape id used in this drawing group.
        /// </summary>
        /// <value>The last MSOSPID.</value>
        public int LastMSOSPID
        {
            get { return field_2_lastMSOSPID; }
            set { field_2_lastMSOSPID = value; }
        }


        /// <summary>
        /// Gets the drawing group id for this record.  This is encoded in the
        /// instance part of the option record.
        /// </summary>
        /// <value>The drawing group id.</value>
        public short DrawingGroupId
        {
            get { return (short)(Options >> 4); }
        }

        /// <summary>
        /// Increments the shape count.
        /// </summary>
        public void IncrementShapeCount()
        {
            this.field_1_numShapes++;
        }
    }
}