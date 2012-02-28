
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
    using System.Collections;
    using NPOI.Util;
    using System.Collections.Generic;

    /// <summary>
    /// This record is used whenever a escher record is encountered that
    /// we do not explicitly support.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    public class UnknownEscherRecord : EscherRecord
    {
        private static byte[] NO_BYTES = new byte[0];

        /** The data for this record not including the the 8 byte header */
        private byte[] thedata = NO_BYTES;
        private List<EscherRecord> childRecords = new List<EscherRecord>();

        public UnknownEscherRecord()
        {
        }

        /// <summary>
        /// This method deSerializes the record from a byte array.
        /// </summary>
        /// <param name="data"> The byte array containing the escher record information</param>
        /// <param name="offset">The starting offset into data </param>
        /// <param name="recordFactory">May be null since this is not a container record.</param>
        /// <returns>The number of bytes Read from the byte array.</returns>
        public override int FillFields(byte[] data, int offset, EscherRecordFactory recordFactory)
        {
            int bytesRemaining = ReadHeader(data, offset);
            /*
		     * Modified by Zhang Zhang
		     * Have a check between avaliable bytes and bytesRemaining, 
		     * take the avaliable length if the bytesRemaining out of range.
		     * July 09, 2010
		     */
            int avaliable = data.Length - (offset + 8);
            if (bytesRemaining > avaliable)
            {
                bytesRemaining = avaliable;
            }
            if (IsContainerRecord)
            {
                int bytesWritten = 0;
                thedata = new byte[0];
                offset += 8;
                bytesWritten += 8;
                while (bytesRemaining > 0)
                {
                    EscherRecord child = recordFactory.CreateRecord(data, offset);
                    int childBytesWritten = child.FillFields(data, offset, recordFactory);
                    bytesWritten += childBytesWritten;
                    offset += childBytesWritten;
                    bytesRemaining -= childBytesWritten;
                    ChildRecords.Add(child);
                }
                return bytesWritten;
            }
            else
            {
                thedata = new byte[bytesRemaining];
                Array.Copy(data, offset + 8, thedata, 0, bytesRemaining);
                return bytesRemaining + 8;
            }
        }

        /// <summary>
        /// Writes this record and any contained records to the supplied byte
        /// array.
        /// </summary>
        /// <param name="offset"></param>
        /// <param name="data"></param>
        /// <param name="listener">a listener for begin and end serialization events.</param>
        /// <returns>the number of bytes written.</returns>
        public override int Serialize(int offset, byte[] data, EscherSerializationListener listener)
        {
            listener.BeforeRecordSerialize(offset, RecordId, this);

            LittleEndian.PutShort(data, offset, Options);
            LittleEndian.PutShort(data, offset + 2, RecordId);
            int remainingBytes = thedata.Length;
            for (IEnumerator iterator = ChildRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)iterator.Current;
                remainingBytes += r.RecordSize;
            }
            LittleEndian.PutInt(data, offset + 4, remainingBytes);
            Array.Copy(thedata, 0, data, offset + 8, thedata.Length);
            int pos = offset + 8 + thedata.Length;
            for (IEnumerator iterator = ChildRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)iterator.Current;
                pos += r.Serialize(pos, data);
            }

            listener.AfterRecordSerialize(pos, RecordId, pos - offset, this);
            return pos - offset;
        }

        /// <summary>
        /// Gets the data.
        /// </summary>
        /// <value>The data.</value>
        public byte[] Data
        {
            get { return thedata; }
        }

        /// <summary>
        /// Returns the number of bytes that are required to Serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + thedata.Length; }
        }

        /// <summary>
        /// Returns the children of this record.  By default this will
        /// be an empty list.  EscherCotainerRecord is the only record
        /// that may contain children.
        /// </summary>
        /// <value></value>
        public override List<EscherRecord> ChildRecords
        {
            get { return childRecords; }
            set { this.childRecords = value; }
        }

        /// <summary>
        /// The short name for this record
        /// </summary>
        /// <value></value>
        public override String RecordName
        {
            get { return "Unknown 0x" + HexDump.ToHex(RecordId); }
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

            StringBuilder children = new StringBuilder();
            if (ChildRecords.Count > 0)
            {
                children.Append("  children: " + nl);
                for (IEnumerator iterator = ChildRecords.GetEnumerator(); iterator.MoveNext(); )
                {
                    EscherRecord record = (EscherRecord)iterator.Current;
                    children.Append(record.ToString());
                    children.Append(nl);
                }
            }

            String theDumpHex = "";
            try
            {
                if (thedata.Length != 0)
                {
                    theDumpHex = "  Extra Data(" + thedata.Length + "):" + nl;
                    theDumpHex += HexDump.Dump(thedata, 0, 0);
                }
            }
            catch (Exception)
            {
                theDumpHex = "Error!!";
            }

            return this.GetType().Name + ":" + nl +
                    "  isContainer: " + IsContainerRecord + nl +
                    "  options: 0x" + HexDump.ToHex(Options) + nl +
                    "  recordId: 0x" + HexDump.ToHex(RecordId) + nl +
                    "  numchildren: " + ChildRecords.Count + nl +
                    theDumpHex +
                    children.ToString();
        }

        /// <summary>
        /// Adds the child record.
        /// </summary>
        /// <param name="childRecord">The child record.</param>
        public void AddChildRecord(EscherRecord childRecord)
        {
            ChildRecords.Add(childRecord);
        }

    }

}