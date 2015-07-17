
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.Record
{

    using NPOI.Util;

    using System;
    using System.IO;


    using NPOI.HSSF.Record.Crypto;
    using System.Diagnostics;


    [Serializable]
    public class LeftoverDataException : Exception
    {
        public LeftoverDataException(int sid, int remainingByteCount)
            : base("Initialisation of record 0x" + StringUtil.ToHexString(sid).ToUpper()
                + " left " + remainingByteCount + " bytes remaining still to be read.")
        {
        }
    }
    internal class SimpleHeaderInput : BiffHeaderInput
    {

        private ILittleEndianInput _lei;

        internal static ILittleEndianInput GetLEI(Stream in1)
        {
            if (in1 is ILittleEndianInput)
            {
                // accessing directly is an optimisation
                return (ILittleEndianInput)in1;
            }
            // less optimal, but should work OK just the same. Often occurs in junit tests.
            return new LittleEndianInputStream(in1);
        }

        public SimpleHeaderInput(Stream in1)
        {
            _lei = GetLEI(in1);
        }
        public int Available()
        {
            return _lei.Available();
        }
        public int ReadDataSize()
        {
            return _lei.ReadUShort();
        }
        public int ReadRecordSID()
        {
            return _lei.ReadUShort();
        }
    }
    /**
     * Title:  Record Input Stream
     * Description:  Wraps a stream and provides helper methods for the construction of records.
     *
     * @author Jason Height (jheight @ apache dot org)
     */

    public class RecordInputStream : Stream, ILittleEndianInput
    {
        /** Maximum size of a single record (minus the 4 byte header) without a continue*/
        public const short MAX_RECORD_DATA_SIZE = 8224;
        private const int INVALID_SID_VALUE = -1;
        private const int DATA_LEN_NEEDS_TO_BE_READ = -1;
        //private const int EOF_RECORD_ENCODED_SIZE = 4;

        //private LittleEndianInput _le;

        protected int _currentSid;
        protected int _currentDataLength = -1;
        protected int _nextSid = -1;
        private int _currentDataOffset = 0;
        // fix warning CS0169 "never used": private long _initialposition;
        private long pos = 0;

        /** Header {@link LittleEndianInput} facet of the wrapped {@link InputStream} */
        private BiffHeaderInput _bhi;
        /** Data {@link LittleEndianInput} facet of the wrapped {@link InputStream} */
        private ILittleEndianInput _dataInput;
        /** the record identifier of the BIFF record currently being read */

        //protected byte[] data = new byte[MAX_RECORD_DATA_SIZE];

        public RecordInputStream(Stream in1)
            : this(in1, null, 0)
        {

        }

        public RecordInputStream(Stream in1, Biff8EncryptionKey key, int initialOffset)
        {
            if (key == null)
            {
                _dataInput = SimpleHeaderInput.GetLEI(in1);
                _bhi = new SimpleHeaderInput(in1);
            }
            else
            {
                Biff8DecryptingStream bds = new Biff8DecryptingStream(in1, initialOffset, key);
                _bhi = bds;
                _dataInput = bds;
            }
            _nextSid = ReadNextSid();
        }

        public int Available()
        {
            return Remaining;
        }

        /** This method will Read a byte from the current record*/
        public int Read()
        {
            CheckRecordPosition(LittleEndianConsts.BYTE_SIZE);
            _currentDataOffset += LittleEndianConsts.BYTE_SIZE;
            pos += LittleEndianConsts.BYTE_SIZE;
            return _dataInput.ReadByte();
        }

        /**
 * 
 * @return the sid of the next record or {@link #INVALID_SID_VALUE} if at end of stream
 */
        private int ReadNextSid()
        {
            int nAvailable = _bhi.Available();
            if (nAvailable < EOFRecord.ENCODED_SIZE)
            {
                if (nAvailable > 0)
                {
                    // some scrap left over?
                    // ex45582-22397.xls has one extra byte after the last record
                    // Excel reads that file OK
                }
                return INVALID_SID_VALUE;
            }
            int result = _bhi.ReadRecordSID();
            if (result == INVALID_SID_VALUE)
            {
                throw new RecordFormatException("Found invalid sid (" + result + ")");
            }
            _currentDataLength = DATA_LEN_NEEDS_TO_BE_READ;

            return result;
        }

        public short Sid
        {
            get { return (short)_currentSid; }
        }

        public override long Position
        {
            get
            {
                return pos;
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public long CurrentLength
        {
            get { return _currentDataLength; }
        }

        public int RecordOffset
        {
            get { return _currentDataOffset; }
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            //if (!this.CanSeek)
            //{
            //    throw new NotSupportedException();
            //}
            //int dwOrigin = 0;
            //switch (origin)
            //{
            //    case SeekOrigin.Begin:
            //        dwOrigin = 0;
            //        if (0L > offset)
            //        {
            //            throw new ArgumentOutOfRangeException("offset", "offset must be positive");
            //        }
            //        this.Position = offset < this.Length ? offset : this.Length;
            //        break;

            //    case SeekOrigin.Current:
            //        dwOrigin = 1;
            //        this.Position = (this.Position + offset) < this.Length ? (this.Position + offset) : this.Length;
            //        break;

            //    case SeekOrigin.End:
            //        dwOrigin = 2;
            //        this.Position = this.Length;
            //        break;

            //    default:
            //        throw new ArgumentException("incorrect SeekOrigin", "origin");
            //}
            //return Position;
            throw new NotSupportedException();
        }

        public bool HasNextRecord
        {
            get
            {
                if (_currentDataLength != -1 && _currentDataLength != _currentDataOffset)
                {
                    throw new LeftoverDataException(_currentSid, Remaining);
                }
                if (_currentDataLength != DATA_LEN_NEEDS_TO_BE_READ)
                {
                    _nextSid = ReadNextSid();
                }
                return _nextSid != INVALID_SID_VALUE;
            }
        }

        /** Moves to the next record in the stream.
         * 
         * <i>Note: The auto continue flag is Reset to true</i>
         */
        public void NextRecord()
        {
            if (_nextSid == INVALID_SID_VALUE)
            {
                throw new InvalidDataException("EOF - next record not available");
            }
            if (_currentDataLength != DATA_LEN_NEEDS_TO_BE_READ)
            {
                throw new InvalidDataException("Cannot call nextRecord() without checking hasNextRecord() first");
            }
            _currentSid = _nextSid;
            _currentDataOffset = 0;
            _currentDataLength = _bhi.ReadDataSize();
            pos += LittleEndianConsts.SHORT_SIZE;
            if (_currentDataLength > MAX_RECORD_DATA_SIZE)
            {
                throw new RecordFormatException("The content of an excel record cannot exceed "
                        + MAX_RECORD_DATA_SIZE + " bytes");
            }
        }

        protected void CheckRecordPosition(int requiredByteCount)
        {
            int nAvailable = Remaining;
            if (nAvailable >= requiredByteCount)
            {
                // all OK
                return;
            }
            if (nAvailable == 0 && IsContinueNext)
            {
                NextRecord();
                return;
            }
            throw new RecordFormatException("Not enough data (" + nAvailable
                    + ") to read requested (" + requiredByteCount + ") bytes");
        }

        /**
         * Reads an 8 bit, signed value
         */
        public override int ReadByte()
        {
            CheckRecordPosition(LittleEndianConsts.BYTE_SIZE);
            _currentDataOffset += LittleEndianConsts.BYTE_SIZE;
            pos += LittleEndianConsts.BYTE_SIZE;
            return _dataInput.ReadByte();
        }

        /**
         * Reads a 16 bit, signed value
         */
        public short ReadShort()
        {
            CheckRecordPosition(LittleEndianConsts.SHORT_SIZE);
            _currentDataOffset += LittleEndianConsts.SHORT_SIZE;
            pos += LittleEndianConsts.SHORT_SIZE;
            return _dataInput.ReadShort();
        }

        public int ReadInt()
        {
            CheckRecordPosition(LittleEndianConsts.INT_SIZE);
            _currentDataOffset += LittleEndianConsts.INT_SIZE;
            pos += LittleEndianConsts.INT_SIZE;
            return _dataInput.ReadInt();
        }

        public long ReadLong()
        {
            CheckRecordPosition(LittleEndianConsts.LONG_SIZE);
            _currentDataOffset += LittleEndianConsts.LONG_SIZE;
            pos += LittleEndianConsts.LONG_SIZE;
            return _dataInput.ReadLong();
        }

        /**
         * Reads an 8 bit, Unsigned value
         */
        public int ReadUByte()
        {
            int s = ReadByte();
            if (s < 0)
            {
                s += 256;
            }
            return s;
        }

        /**
         * Reads a 16 bit,un- signed value.
         * @return
         */
        public int ReadUShort()
        {
            CheckRecordPosition(LittleEndianConsts.SHORT_SIZE);
            _currentDataOffset += LittleEndianConsts.SHORT_SIZE;
            pos += LittleEndianConsts.SHORT_SIZE;
            return _dataInput.ReadUShort();
        }

        public double ReadDouble()
        {
            CheckRecordPosition(LittleEndianConsts.DOUBLE_SIZE);
            _currentDataOffset += LittleEndianConsts.DOUBLE_SIZE;

            long valueLongBits = _dataInput.ReadLong();
            double result = BitConverter.Int64BitsToDouble(valueLongBits);
            //Excel represents NAN in several ways, at this point in time we do not often
            //know the sequence of bytes, so as a hack we store the NAN byte sequence
            //so that it is not corrupted.
            //if (double.IsNaN(result))
            //{
            //throw new Exception("Did not expect to read NaN"); // (Because Excel typically doesn't write NaN
            //}
            pos += LittleEndianConsts.DOUBLE_SIZE;
            return result;
        }
        public void ReadFully(byte[] buf)
        {
            ReadFully(buf, 0, buf.Length);
        }

        public void ReadFully(byte[] buf, int off, int len)
        {
            /*CheckRecordPosition(len);
            _dataInput.ReadFully(buf, off, len);
            _currentDataOffset += len;
            pos += len;*/

            int origLen = len;
            if (buf == null)
            {
                throw new ArgumentNullException();
            }
            else if (off < 0 || len < 0 || len > buf.Length - off)
            {
                throw new IndexOutOfRangeException();
            }

            while (len > 0)
            {
                int nextChunk = Math.Min(Available(), len);
                if (nextChunk == 0)
                {
                    if (!HasNextRecord)
                    {
                        throw new RecordFormatException("Can't read the remaining " + len + " bytes of the requested " + origLen + " bytes. No further record exists.");
                    }
                    else
                    {
                        NextRecord();
                        nextChunk = Math.Min(Available(), len);
                        Debug.Assert(nextChunk > 0);
                    }
                }
                CheckRecordPosition(nextChunk);
                _dataInput.ReadFully(buf, off, nextChunk);
                _currentDataOffset += nextChunk;
                off += nextChunk;
                len -= nextChunk;

                pos += nextChunk;
            }
        }
        /**     
         *  given a byte array of 16-bit Unicode Chars, compress to 8-bit and     
         *  return a string     
         *     
         * { 0x16, 0x00 } -0x16     
         *      
         * @param Length the Length of the string
         * @return                                     the Converted string
         * @exception  ArgumentException        if len is too large (i.e.,
         *      there is not enough data in string to Create a String of that     
         *      Length)     
         */
        public String ReadUnicodeLEString(int requestedLength)
        {
            return ReadStringCommon(requestedLength, false);
        }

        public String ReadCompressedUnicode(int requestedLength)
        {
            return ReadStringCommon(requestedLength, true);
        }
        private String ReadStringCommon(int requestedLength, bool pIsCompressedEncoding)
        {
            // Sanity check to detect garbage string lengths
            if (requestedLength < 0 || requestedLength > 0x100000)
            { // 16 million chars?
                throw new ArgumentException("Bad requested string length (" + requestedLength + ")");
            }
            char[] buf = new char[requestedLength];
            bool isCompressedEncoding = pIsCompressedEncoding;
            int curLen = 0;
            while (true)
            {
                int availableChars = isCompressedEncoding ? Remaining : Remaining / LittleEndianConsts.SHORT_SIZE;
                if (requestedLength - curLen <= availableChars)
                {
                    // enough space in current record, so just read it out
                    while (curLen < requestedLength)
                    {
                        char ch;
                        if (isCompressedEncoding)
                        {
                            ch = (char)ReadUByte();
                        }
                        else
                        {
                            ch = (char)ReadShort();
                        }
                        buf[curLen] = ch;
                        curLen++;
                    }
                    return new String(buf);// Encoding.UTF8.GetChars(buf,0,buf.Length);
                }
                // else string has been spilled into next continue record
                // so read what's left of the current record
                while (availableChars > 0)
                {
                    char ch;
                    if (isCompressedEncoding)
                    {
                        ch = (char)ReadUByte();
                    }
                    else
                    {
                        ch = (char)ReadShort();
                    }
                    buf[curLen] = ch;
                    curLen++;
                    availableChars--;
                }
                if (!IsContinueNext)
                {
                    throw new RecordFormatException("Expected to find a ContinueRecord in order to read remaining "
                            + (requestedLength - curLen) + " of " + requestedLength + " chars");
                }
                if (Remaining != 0)
                {
                    throw new RecordFormatException("Odd number of bytes(" + Remaining + ") left behind");
                }
                NextRecord();
                // note - the compressed flag may change on the fly
                byte compressFlag = (byte)ReadByte();
                Debug.Assert(compressFlag == 0 || compressFlag == 1);
                isCompressedEncoding = (compressFlag == 0);
            }
        }
        public String ReadString()
        {
            int requestedLength = ReadUShort();
            byte compressFlag = (byte)ReadByte();
            return ReadStringCommon(requestedLength, compressFlag == 0);
        }

        /** Returns the remaining bytes for the current record.
         * 
         * @return The remaining bytes of the current record.
         */
        public byte[] ReadRemainder()
        {
            int size = Remaining;
            if (size == 0)
            {
                return new byte[0];
            }
            byte[] result = new byte[size];
            ReadFully(result);
            return result;
        }

        /** Reads all byte data for the current record, including any
         *  that overlaps into any following continue records.
         * 
         *  @deprecated Best to write a input stream that wraps this one where there Is
         *  special sub record that may overlap continue records.
         */
        public byte[] ReadAllContinuedRemainder()
        {
            //Using a ByteArrayOutputStream is just an easy way to Get a
            //growable array of the data.
            using (MemoryStream out1 = new MemoryStream(2 * MAX_RECORD_DATA_SIZE))
            {

                while (true)
                {
                    byte[] b = ReadRemainder();
                    out1.Write(b, 0, b.Length);
                    if (!IsContinueNext)
                    {
                        break;
                    }
                    NextRecord();
                }

                return out1.ToArray();
            }
        }

        /** The remaining number of bytes in the <i>current</i> record.
         * 
         * @return The number of bytes remaining in the current record
         */
        public int Remaining
        {
            get
            {
                if (_currentDataLength == DATA_LEN_NEEDS_TO_BE_READ)
                {
                    // already read sid of next record. so current one is finished
                    return 0;
                }
                return _currentDataLength - _currentDataOffset;
            }
        }

        /** Returns true iif a Continue record is next in the excel stream _currentDataOffset
         * 
         * @return True when a ContinueRecord is next.
         */
        public bool IsContinueNext
        {
            get
            {
                if (_currentDataLength != DATA_LEN_NEEDS_TO_BE_READ && _currentDataOffset != _currentDataLength)
                {
                    throw new InvalidOperationException("Should never be called before end of current record");
                }
                if (!HasNextRecord)
                {
                    return false;
                }
                // At what point are records continued?
                //  - Often from within the char data of long strings (caller is within readStringCommon()).
                //  - From UnicodeString construction (many different points - call via checkRecordPosition)
                //  - During TextObjectRecord construction (just before the text, perhaps within the text, 
                //    and before the formatting run data)
                return _nextSid == ContinueRecord.sid;

            }
        }

        public override long Length
        {
            get { return _currentDataLength; }
        }

        public override void SetLength(long value)
        {
            _currentDataLength = (int)value;
        }

        public override void Flush()
        {
            throw new NotSupportedException();
        }

        // Properties
        public override bool CanRead
        {
            get
            {
                return true;
            }
        }

        public override bool CanSeek
        {
            get
            {
                return false;
            }
        }

        public override bool CanWrite
        {
            get
            {
                return false;
            }
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        public override int Read(byte[] b, int off, int len)
        {
            int limit = Math.Min(len, Remaining);
            if (limit == 0)
            {
                return 0;
            }
            ReadFully(b, off, limit);
            return limit;
            //Array.Copy(data, _currentDataOffset, b, off, len);
            //_currentDataOffset += len;
            //return Math.Min(data.Length, b.Length);
        }


        /**
 @return sid of next record. Can be called after hasNextRecord()
 */
        public int GetNextSid()
        {
            return _nextSid;
        }
    }
}