/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace NPOI.HSSF.Record
{
    using System;
    using NPOI.Util;


    /**
     * Subclasses of this class (the majority of BIFF records) are non-continuable.  This allows for
     * some simplification of serialization logic
     * 
     * @author Josh Micich
     */
    public abstract class StandardRecord : Record
    {

        protected abstract int DataSize{get;}
        public override int RecordSize
        {
            get
            {
                return 4 + DataSize;
            }
        }

        /// <summary>
        /// Write the data content of this BIFF record including the sid and record length.
        /// The subclass must write the exact number of bytes as reported by Record#getRecordSize()
        /// </summary>
        /// <param name="offset">offset</param>
        /// <param name="data">data</param>
        /// <returns></returns>
        public override int Serialize(int offset, byte[] data)
        {
            int dataSize = DataSize;
            int recSize = 4 + dataSize;
            LittleEndianByteArrayOutputStream out1 = new LittleEndianByteArrayOutputStream(data, offset, recSize);
            out1.WriteShort(this.Sid);
            out1.WriteShort(dataSize);
            Serialize(out1);
            if (out1.WriteIndex - offset != recSize)
            {
                throw new InvalidOperationException("Error in serialization of (" + this.GetType().Name + "): "
                        + "Incorrect number of bytes written - expected "
                        + recSize + " but got " + (out1.WriteIndex - offset));
            }
            return recSize;
        }

        /**
         * Write the data content of this BIFF record.  The 'ushort sid' and 'ushort size' header fields
         * have already been written by the superclass.<br/>
         * 
         * The number of bytes written must equal the record size reported by
         * {@link Record#getDataSize()} minus four
         * ( record header consiting of a 'ushort sid' and 'ushort reclength' has already been written
         * by thye superclass).
         */
        public abstract void Serialize(ILittleEndianOutput out1);
    }
}