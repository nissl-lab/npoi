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
     * An augmented {@link LittleEndianOutput} used for serialization of {@link ContinuableRecord}s.
     * This class keeps track of how much remaining space is available in the current BIFF record and
     * can start new {@link ContinueRecord}s as required. 
     * 
     * @author Josh Micich
     */
    public class ContinuableRecordOutput : ILittleEndianOutput
    {

        private ILittleEndianOutput _out;
        private UnknownLengthRecordOutput _ulrOutput;
        private int _totalPreviousRecordsSize;

        internal ContinuableRecordOutput(ILittleEndianOutput out1, int sid)
        {
            _ulrOutput = new UnknownLengthRecordOutput(out1, sid);
            _out = out1;
            _totalPreviousRecordsSize = 0;
        }

        public static ContinuableRecordOutput CreateForCountingOnly()
        {
            return new ContinuableRecordOutput(NOPOutput, -777); // fake sid
        }

        /**
         * @return total number of bytes written so far (including all BIFF headers) 
         */
        public int TotalSize
        {
            get
            {
                return _totalPreviousRecordsSize + _ulrOutput.TotalSize;
            }
        }
        /**
         * Terminates the last record (also updates its 'ushort size' field)
         */
        public void Terminate()
        {
            _ulrOutput.Terminate();
        }
        /**
         * @return number of remaining bytes of space in current record
         */
        public int AvailableSpace
        {
            get
            {
                return _ulrOutput.AvailableSpace;
            }
        }

        /**
         * Terminates the current record and starts a new {@link ContinueRecord} (regardless
         * of how much space is still available in the current record).
         */
        public void WriteContinue()
        {
            _ulrOutput.Terminate();
            _totalPreviousRecordsSize += _ulrOutput.TotalSize;
            _ulrOutput = new UnknownLengthRecordOutput(_out, ContinueRecord.sid);
        }
        public void WriteContinueIfRequired(int requiredContinuousSize)
        {
            if (_ulrOutput.AvailableSpace < requiredContinuousSize)
            {
                WriteContinue();
            }
        }

        /**
         * Writes the 'optionFlags' byte and encoded character data of a unicode string.  This includes:
         * <ul>
         * <li>byte optionFlags</li>
         * <li>encoded character data (in "ISO-8859-1" or "UTF-16LE" encoding)</li>
         * </ul>
         * 
         * Notes:
         * <ul>
         * <li>The value of the 'is16bitEncoded' flag is determined by the actual character data 
         * of <c>text</c></li>
         * <li>The string options flag is never separated (by a {@link ContinueRecord}) from the
         * first chunk of character data it refers to.</li>
         * <li>The 'ushort Length' field is assumed to have been explicitly written earlier.  Hence, 
         * there may be an intervening {@link ContinueRecord}</li>
         * </ul>
         */
        public void WriteStringData(String text)
        {
            bool is16bitEncoded = StringUtil.HasMultibyte(text);
            // calculate total size of the header and first encoded char
            int keepTogetherSize = 1 + 1; // ushort len, at least one character byte
            int optionFlags = 0x00;
            if (is16bitEncoded)
            {
                optionFlags |= 0x01;
                keepTogetherSize += 1; // one extra byte for first char
            }
            WriteContinueIfRequired(keepTogetherSize);
            WriteByte(optionFlags);
            WriteCharacterData(text, is16bitEncoded);
        }
        /**
         * Writes a unicode string complete with header and character data.  This includes:
         * <ul>
         * <li>ushort Length</li>
         * <li>byte optionFlags</li>
         * <li>ushort numberOfRichTextRuns (optional)</li>
         * <li>ushort extendedDataSize (optional)</li>
         * <li>encoded character data (in "ISO-8859-1" or "UTF-16LE" encoding)</li>
         * </ul>
         * 
         * The following bits of the 'optionFlags' byte will be set as appropriate:
         * <table border='1'>
         * <tr><th>Mask</th><th>Description</th></tr>
         * <tr><td>0x01</td><td>is16bitEncoded</td></tr>
         * <tr><td>0x04</td><td>hasExtendedData</td></tr>
         * <tr><td>0x08</td><td>isRichText</td></tr>
         * </table>
         * Notes:
         * <ul> 
         * <li>The value of the 'is16bitEncoded' flag is determined by the actual character data 
         * of <c>text</c></li>
         * <li>The string header fields are never separated (by a {@link ContinueRecord}) from the
         * first chunk of character data (i.e. the first character is always encoded in the same
         * record as the string header).</li>
         * </ul>
         */
        public void WriteString(String text, int numberOfRichTextRuns, int extendedDataSize)
        {
            bool is16bitEncoded = StringUtil.HasMultibyte(text);
            // calculate total size of the header and first encoded char
            int keepTogetherSize = 2 + 1 + 1; // ushort len, byte optionFlags, at least one character byte
            int optionFlags = 0x00;
            if (is16bitEncoded)
            {
                optionFlags |= 0x01;
                keepTogetherSize += 1; // one extra byte for first char
            }
            if (numberOfRichTextRuns > 0)
            {
                optionFlags |= 0x08;
                keepTogetherSize += 2;
            }
            if (extendedDataSize > 0)
            {
                optionFlags |= 0x04;
                keepTogetherSize += 4;
            }
            WriteContinueIfRequired(keepTogetherSize);
            WriteShort(text.Length);
            WriteByte(optionFlags);
            if (numberOfRichTextRuns > 0)
            {
                WriteShort(numberOfRichTextRuns);
            }
            if (extendedDataSize > 0)
            {
                WriteInt(extendedDataSize);
            }
            WriteCharacterData(text, is16bitEncoded);
        }


        private void WriteCharacterData(String text, bool is16bitEncoded)
        {
            int nChars = text.Length;
            int i = 0;
            if (is16bitEncoded)
            {
                while (true)
                {
                    int nWritableChars = Math.Min(nChars - i, _ulrOutput.AvailableSpace / 2);
                    for (; nWritableChars > 0; nWritableChars--)
                    {
                        _ulrOutput.WriteShort(text[i++]);
                    }
                    if (i >= nChars)
                    {
                        break;
                    }
                    WriteContinue();
                    WriteByte(0x01);
                }
            }
            else
            {
                while (true)
                {
                    int nWritableChars = Math.Min(nChars - i, _ulrOutput.AvailableSpace / 1);
                    for (; nWritableChars > 0; nWritableChars--)
                    {
                        _ulrOutput.WriteByte(text[i++]);
                    }
                    if (i >= nChars)
                    {
                        break;
                    }
                    WriteContinue();
                    WriteByte(0x00);
                }
            }
        }

        public void Write(byte[] b)
        {
            WriteContinueIfRequired(b.Length);
            _ulrOutput.Write(b);
        }
        public void Write(byte[] b, int offset, int len)
        {
            //WriteContinueIfRequired(len);
            //_ulrOutput.Write(b, offset, len);
            int i = 0;
            while (true)
            {
                int nWritableChars = Math.Min(len - i, _ulrOutput.AvailableSpace / 1);
                for (; nWritableChars > 0; nWritableChars--)
                {
                    _ulrOutput.WriteByte(b[offset + i++]);
                }
                if (i >= len)
                {
                    break;
                }
                WriteContinue();
            }
        }
        public void WriteByte(int v)
        {
            WriteContinueIfRequired(1);
            _ulrOutput.WriteByte(v);
        }
        public void WriteDouble(double v)
        {
            WriteContinueIfRequired(8);
            _ulrOutput.WriteDouble(v);
        }
        public void WriteInt(int v)
        {
            WriteContinueIfRequired(4);
            _ulrOutput.WriteInt(v);
        }
        public void WriteLong(long v)
        {
            WriteContinueIfRequired(8);
            _ulrOutput.WriteLong(v);
        }
        public void WriteShort(int v)
        {
            WriteContinueIfRequired(2);
            _ulrOutput.WriteShort(v);
        }

        ///**
        // * Allows optimised usage of {@link ContinuableRecordOutput} for sizing purposes only.
        // */
        private static ILittleEndianOutput NOPOutput = new DelayableLittleEndianOutput1();

        class DelayableLittleEndianOutput1 : IDelayableLittleEndianOutput
        {

            public ILittleEndianOutput CreateDelayedOutput(int size)
            {
                return this;
            }
            public void Write(byte[] b)
            {
                // does nothing
            }
            public void Write(byte[] b, int offset, int len)
            {
                // does nothing
            }
            public void WriteByte(int v)
            {
                // does nothing
            }
            public void WriteDouble(double v)
            {
                // does nothing
            }
            public void WriteInt(int v)
            {
                // does nothing
            }
            public void WriteLong(long v)
            {
                // does nothing
            }
            public void WriteShort(int v)
            {
                // does nothing
            }
        }
    }
}
