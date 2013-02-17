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
    using System;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    /**
     * Allows the writing of BIFF records when the 'ushort size' header field is not known in advance.
     * When the client is finished writing data, it calls {@link #terminate()}, at which point this 
     * class updates the 'ushort size' with its value. 
     * 
     * @author Josh Micich
     */
    class UnknownLengthRecordOutput : ILittleEndianOutput
    {
        private const int MAX_DATA_SIZE = RecordInputStream.MAX_RECORD_DATA_SIZE;

        private ILittleEndianOutput _originalOut;
        /** for writing the 'ushort size'  field once its value is known */
        private ILittleEndianOutput _dataSizeOutput;
        private byte[] _byteBuffer;
        private ILittleEndianOutput _out;
        private int _size;

        public UnknownLengthRecordOutput(ILittleEndianOutput out1, int sid)
        {
            _originalOut = out1;
            out1.WriteShort(sid);
            if (out1 is IDelayableLittleEndianOutput)
            {
                // optimisation
                IDelayableLittleEndianOutput dleo = (IDelayableLittleEndianOutput)out1;
                _dataSizeOutput = dleo.CreateDelayedOutput(2);
                _byteBuffer = null;
                _out = out1;
            }
            else
            {
                // otherwise temporarily Write all subsequent data to a buffer
                _dataSizeOutput = out1;
                _byteBuffer = new byte[RecordInputStream.MAX_RECORD_DATA_SIZE];
                _out = new LittleEndianByteArrayOutputStream(_byteBuffer, 0);
            }
        }
        /**
         * includes 4 byte header
         */
        public int TotalSize
        {
            get
            {
                return 4 + _size;
            }
        }
        public int AvailableSpace
        {
            get
            {
                if (_out == null)
                {
                    throw new InvalidOperationException("Record already terminated");
                }
                return MAX_DATA_SIZE - _size;
            }
        }
        /**
         * Finishes writing the current record and updates 'ushort size' field.<br/>
         * After this method is called, only {@link #getTotalSize()} may be called.
         */
        public void Terminate()
        {
            if (_out == null)
            {
                throw new InvalidOperationException("Record already terminated");
            }
            _dataSizeOutput.WriteShort(_size);
            if (_byteBuffer != null)
            {
                _originalOut.Write(_byteBuffer, 0, _size);
                _out = null;
                return;
            }
            _out = null;
        }

        public void Write(byte[] b)
        {
            _out.Write(b);
            _size += b.Length;
        }
        public void Write(byte[] b, int offset, int len)
        {
            _out.Write(b, offset, len);
            _size += len;
        }
        public void WriteByte(int v)
        {
            _out.WriteByte(v);
            _size += 1;
        }
        public void WriteDouble(double v)
        {
            _out.WriteDouble(v);
            _size += 8;
        }
        public void WriteInt(int v)
        {
            _out.WriteInt(v);
            _size += 4;
        }
        public void WriteLong(long v)
        {
            _out.WriteLong(v);
            _size += 8;
        }
        public void WriteShort(int v)
        {
            _out.WriteShort(v);
            _size += 2;
        }
    }
}