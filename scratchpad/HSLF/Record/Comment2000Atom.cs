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

namespace NPOI.HSLF.Record
{
    using System;
    using System.IO;
    using NPOI.Util;
    using NPOI.HSLF.Util;

    /**
     * An atomic record Containing information about a comment.
     *
     * @author Daniel Noll
     */

    public class Comment2000Atom : RecordAtom
    {
        /**
         * Record header.
         */
        private byte[] _header;

        /**
         * Record data.
         */
        private byte[] _data;

        /**
         * Constructs a brand new comment atom record.
         */
        public Comment2000Atom()
        {
            _header = new byte[8];
            _data = new byte[28];

            LittleEndian.PutShort(_header, 2, (short)RecordType);
            LittleEndian.PutInt(_header, 4, _data.Length);

            // It is fine for the other values to be zero
        }

        /**
         * Constructs the comment atom record from its source data.
         *
         * @param source the source data as a byte array.
         * @param start the start offset into the byte array.
         * @param len the length of the slice in the byte array.
         */
        public Comment2000Atom(byte[] source, int start, int len)
        {
            // Get the header.
            _header = new byte[8];
            Array.Copy(source, start, _header, 0, 8);

            // Get the record data.
            _data = new byte[len - 8];
            Array.Copy(source, start + 8, _data, 0, len - 8);
        }

        /**
         * Gets the comment number (note - each user normally has their own count).
         * @return the comment number.
         */
        public int Number
        {
            get
            {
                return LittleEndian.GetInt(_data, 0);
            }
            set
            {
                LittleEndian.PutInt(_data, 0, value);
            }
        }
        /**
         * Gets the date the comment was made.
         * @return the comment date.
         */
        public DateTime Date
        {
            get
            {
                return SystemTimeUtils.GetDate(_data, 4);
            }
            set
            {
                SystemTimeUtils.StoreDate(value, _data, 4);
            }
        }
        /**
         * Gets the X offset of the comment on the page.
         * @return the X OffSet.
         */
        public int XOffset
        {
            get
            {
                return LittleEndian.GetInt(_data, 20);
            }
            set
            {
                LittleEndian.PutInt(_data, 20, value);
            }
        }
        /**
         * Gets the Y offset of the comment on the page.
         * @return the Y OffSet.
         */
        public int YOffset
        {
            get
            {
                return LittleEndian.GetInt(_data, 24);
            }
            set
            {
                LittleEndian.PutInt(_data, 24, value);
            }
        }
        /**
         * Gets the record type.
         * @return the record type.
         */
        public override long RecordType
        {
            get
            {
                return RecordTypes.Comment2000Atom.typeID;
            }
        }

        /**
         * Write the contents of the record back, so it can be written
         * to disk
         *
         * @param out the output stream to write to.
         * @throws IOException if an error occurs.
         */
        public override void WriteOut(Stream out1)
        {
            out1.Write(_header,(int)out1.Position,_header.Length);
            out1.Write(_data,(int)out1.Position,_data.Length);
        }
    }
}

