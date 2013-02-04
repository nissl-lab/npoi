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

namespace NPOI.HSSF.Record.Cont
{
    using NPOI.HSSF.Record;
    using NPOI.Util;

    /**
     * Common superclass of all records that can produce {@link ContinueRecord}s while being Serialized.
     * 
     * @author Josh Micich
     */
    public abstract class ContinuableRecord : Record
    {

        protected ContinuableRecord()
        {
            // no fields to initialise 
        }
        /**
         * Serializes this record's content to the supplied data output.<br/>
         * The standard BIFF header (ushort sid, ushort size) has been handled by the superclass, so 
         * only BIFF data should be written by this method.  Simple data types can be written with the
         * standard {@link LittleEndianOutput} methods.  Methods from {@link ContinuableRecordOutput} 
         * can be used to Serialize strings (with {@link ContinueRecord}s being written as required).
         * If necessary, implementors can explicitly start {@link ContinueRecord}s (regardless of the
         * amount of remaining space).
         * 
         * @param out a data output stream
         */
        protected abstract void Serialize(ContinuableRecordOutput out1);


        /**
         * @return the total Length of the encoded record(s) 
         * (Note - if any {@link ContinueRecord} is required, this result includes the
         * size of those too)
         */
        public override int RecordSize
        {
            get
            {
                ContinuableRecordOutput out1 = ContinuableRecordOutput.CreateForCountingOnly();
                Serialize(out1);
                out1.Terminate();
                return out1.TotalSize;
            }
        }

        public override int Serialize(int offset, byte[] data)
        {
            ILittleEndianOutput leo = new LittleEndianByteArrayOutputStream(data, offset);
            ContinuableRecordOutput out1 = new ContinuableRecordOutput(leo, this.Sid);
            Serialize(out1);
            out1.Terminate();
            return out1.TotalSize;
        }
    }
}