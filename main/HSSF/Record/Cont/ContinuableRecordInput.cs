/*
 *  ====================================================================
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for Additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * ====================================================================
 */

namespace NPOI.HSSF.Record.Cont
{

    using NPOI.HSSF.Record;
    using NPOI.Util;

    /**
     * A decorated {@link RecordInputStream} that can read primitive data types
     * (short, int, long, etc.) spanned across a {@link ContinueRecord } boundary.
     *
     * <p>
     *  Most records construct themselves from {@link RecordInputStream}.
     *  This class assumes that a {@link ContinueRecord} record break always occurs at the type boundary,
     *  however, it is not always so.
     * </p>
     *  Two  attachments to <a href="https://issues.apache.org/bugzilla/Show_bug.cgi?id=50779">Bugzilla 50779</a>
     *  demonstrate that a CONTINUE break can appear right in between two bytes of a unicode character
     *  or between two bytes of a <code>short</code>. The problematic portion of the data is
     *  in a Asian Phonetic Settings Block (ExtRst) of a UnicodeString.
     * <p>
     *  {@link RecordInputStream} greedily requests the bytes to be read and stumbles on such files with a
     *  "Not enough data (1) to read requested (2) bytes" exception.  The <code>ContinuableRecordInput</code>
     *   class circumvents this "type boundary" rule and Reads data byte-by-byte rolling over CONTINUE if necessary.
     * </p>
     *
     * <p>
     * YK: For now (March 2011) this class is only used to read
     *   @link NPOI.HSSF.Record.Common.UnicodeString.ExtRst} blocks of a UnicodeString.
     *
     * </p>
     *
     * @author Yegor Kozlov
     */
    public class ContinuableRecordInput : ILittleEndianInput
    {
        private RecordInputStream _in;

        public ContinuableRecordInput(RecordInputStream in1)
        {
            _in = in1;
        }
        public int Available()
        {
            return _in.Available();
        }

        public int ReadByte()
        {
            return _in.ReadByte();
        }

        public int ReadUByte()
        {
            return _in.ReadUByte();
        }

        public short ReadShort()
        {
            return _in.ReadShort();
        }

        public int ReadUShort()
        {
            int ch1 = ReadUByte();
            int ch2 = ReadUByte();
            return (ch2 << 8) + (ch1 << 0);
        }

        public int ReadInt()
        {
            int ch1 = _in.ReadUByte();
            int ch2 = _in.ReadUByte();
            int ch3 = _in.ReadUByte();
            int ch4 = _in.ReadUByte();
            return (ch4 << 24) + (ch3 << 16) + (ch2 << 8) + (ch1 << 0);
        }

        public long ReadLong()
        {
            int b0 = _in.ReadUByte();
            int b1 = _in.ReadUByte();
            int b2 = _in.ReadUByte();
            int b3 = _in.ReadUByte();
            int b4 = _in.ReadUByte();
            int b5 = _in.ReadUByte();
            int b6 = _in.ReadUByte();
            int b7 = _in.ReadUByte();
            return (((long)b7 << 56) +
                    ((long)b6 << 48) +
                    ((long)b5 << 40) +
                    ((long)b4 << 32) +
                    ((long)b3 << 24) +
                    (b2 << 16) +
                    (b1 << 8) +
                    (b0 << 0));
        }

        public double ReadDouble()
        {
            return _in.ReadDouble();
        }
        public void ReadFully(byte[] buf)
        {
            _in.ReadFully(buf);
        }
        public void ReadFully(byte[] buf, int off, int len)
        {
            _in.ReadFully(buf, off, len);
        }

    }

}