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
    using NPOI.Util;
    using System.IO;
    using System.Collections;

    /**
     * If we come across a record we know has children of (potential)
     *  interest, but where the record itself is boring, but where other
     *  records may care about where this one lives, we create one
     *  of these. It allows us to get at the children, and track where on
     *  disk this is, but not much else.
     * Anything done using this should quite quickly be transitioned to its
     *  own proper record class!
     *
     * @author Nick Burch
     */

    public class DummyPositionSensitiveRecordWithChildren : PositionDependentRecordContainer
    {
        private byte[] _header;
        private long _type;

        /**
         * Create a new holder for a boring record with children, but with
         *  position dependent characteristics
         */
        protected DummyPositionSensitiveRecordWithChildren(byte[] source, int start, int len)
        {
            // Just grab the header, not the whole contents
            _header = new byte[8];
            Array.Copy(source, start, _header, 0, 8);
            _type = LittleEndian.GetUShort(_header, 2);

            // Find our children
            _children = Record.FindChildRecords(source, start + 8, len - 8);
        }

        /**
         * Return the value we were given at creation
         */
        public override long RecordType
        {
            get { return _type; }
        }

        /**
         * Write the contents of the record back, so it can be written
         *  to disk
         */
        public override void WriteOut(Stream out1)
        {
            WriteOut(_header[0], _header[1], _type, _children, out1);
        }
    }
}


