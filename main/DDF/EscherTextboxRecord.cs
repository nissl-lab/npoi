
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
    using NPOI.HSSF.Record;


    /// <summary>
    /// Holds data from the parent application. Most commonly used to store
    /// text in the format of the parent application, rather than in
    /// Escher format. We don't attempt to understand the contents, since
    /// they will be in the parent's format, not Escher format.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// @author Nick Burch  (nick at torchbox dot com)
    /// </summary>
    public class EscherTextboxRecord : EscherRecord
    {
        public static short RECORD_ID = unchecked((short)0xF00D);
        public static String RECORD_DESCRIPTION = "msofbtClientTextbox";

        private static byte[] NO_BYTES = new byte[0];

        /** The data for this record not including the the 8 byte header */
        private byte[] thedata = NO_BYTES;

        public EscherTextboxRecord()
        {
        }

        /**
         * This method deserializes the record from a byte array.
         *
         * @param data          The byte array containing the escher record information
         * @param offset        The starting offset into <c>data</c>.
         * @param recordFactory May be null since this is not a container record.
         * @return The number of bytes Read from the byte array.
         */
        public override int FillFields(byte[] data, int offset, EscherRecordFactory recordFactory)
        {
            int bytesRemaining = ReadHeader(data, offset);

            // Save the data, Ready for the calling code to do something
            //  useful with it
            thedata = new byte[bytesRemaining];
            Array.Copy(data, offset + 8, thedata, 0, bytesRemaining);
            return bytesRemaining + 8;
        }

        /// <summary>
        /// Writes this record and any contained records to the supplied byte
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
            LittleEndian.PutInt(data, offset + 4, remainingBytes);
            Array.Copy(thedata, 0, data, offset + 8, thedata.Length);
            int pos = offset + 8 + thedata.Length;

            listener.AfterRecordSerialize(pos, RecordId, pos - offset, this);
            int size = pos - offset;
            if (size != RecordSize)
                throw new RecordFormatException(size + " bytes written but RecordSize reports " + RecordSize);
            return size;
        }

        /// <summary>
        /// Returns any extra data associated with this record.  In practice excel
        /// does not seem to put anything here, but with PowerPoint this will
        /// contain the bytes that make up a TextHeaderAtom followed by a
        /// TextBytesAtom/TextCharsAtom
        /// </summary>
        /// <value>The data.</value>
        public byte[] Data
        {
            get { return thedata; }
        }

        /// <summary>
        /// Sets the extra data (in the parent application's format) to be
        /// contained by the record. Used when the parent application changes
        /// the contents.
        /// </summary>
        /// <param name="b">The b.</param>
        /// <param name="start">The start.</param>
        /// <param name="Length">The length.</param>
        public void SetData(byte[] b, int start, int Length)
        {
            thedata = new byte[Length];
            Array.Copy(b, start, thedata, 0, Length);
        }
        /// <summary>
        /// Sets the data.
        /// </summary>
        /// <param name="b">The b.</param>
        public void SetData(byte[] b)
        {
            SetData(b, 0, b.Length);
        }


        /// <summary>
        /// Returns the number of bytes that are required to serialize this record.
        /// </summary>
        /// <value>Number of bytes</value>
        public override int RecordSize
        {
            get { return 8 + thedata.Length; }
        }

        /// <summary>
        /// The short name for this record
        /// </summary>
        /// <value></value>
        public override String RecordName
        {
            get { return "ClientTextbox"; }
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

            String theDumpHex = "";
            try
            {
                if (thedata.Length != 0)
                {
                    theDumpHex = "  Extra Data:" + nl;
                    theDumpHex += HexDump.Dump(thedata, 0, 0);
                }
            }
            catch (Exception)
            {
                theDumpHex = "Error!!";
            }

            return GetType().Name + ":" + nl +
                    "  isContainer: " + IsContainerRecord + nl +
                    "  options: 0x" + HexDump.ToHex(Options) + nl +
                    "  recordId: 0x" + HexDump.ToHex(RecordId) + nl +
                    "  numchildren: " + ChildRecords.Count + nl +
                    theDumpHex;
        }

    }

}

