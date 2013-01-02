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

using System;
namespace NPOI.POIFS.Storage
{

    /**
     * Wraps a <c>byte</c> array and provides simple data input access.
     * Internally, this class maintains a buffer read index, so that for the most part, primitive
     * data can be read in a data-input-stream-like manner.<p/>
     *
     * Note - the calling class should call the {@link #available()} method to detect end-of-buffer
     * and Move to the next data block when the current is exhausted.
     * For optimisation reasons, no error handling is performed in this class.  Thus, mistakes in
     * calling code ran may raise ugly exceptions here, like {@link ArrayIndexOutOfBoundsException},
     * etc .<p/>
     *
     * The multi-byte primitive input methods ({@link #readUshortLE()}, {@link #readIntLE()} and
     * {@link #readLongLE()}) have corresponding 'spanning Read' methods which (when required) perform
     * a read across the block boundary.  These spanning read methods take the previous
     * {@link DataInputBlock} as a parameter.
     * Reads of larger amounts of data (into <c>byte</c> array buffers) must be managed by the caller
     * since these could conceivably involve more than two blocks.
     *
     * @author Josh Micich
     */
    public class DataInputBlock
    {

        /**
         * Possibly any size (usually 512K or 64K).  Assumed to be at least 8 bytes for all blocks
         * before the end of the stream.  The last block in the stream can be any size except zero. 
         */
        private byte[] _buf;
        private int _readIndex;
        private int _maxIndex;

        internal DataInputBlock(byte[] data, int startOffset)
        {
            _buf = data;
            _readIndex = startOffset;
            _maxIndex = _buf.Length;
        }
        public int Available()
        {
            return _maxIndex - _readIndex;
        }

        public int ReadUByte()
        {
            return _buf[_readIndex++] & 0xFF;
        }

        /**
         * Reads a <c>short</c> which was encoded in <em>little endian</em> format.
         */
        public int ReadUshortLE()
        {
            int i = _readIndex;

            int b0 = _buf[i++] & 0xFF;
            int b1 = _buf[i++] & 0xFF;
            _readIndex = i;
            return (b1 << 8) + (b0 << 0);
        }

        /**
         * Reads a <c>short</c> which spans the end of <c>prevBlock</c> and the start of this block.
         */
        public int ReadUshortLE(DataInputBlock prevBlock)
        {
            // simple case - will always be one byte in each block
            int i = prevBlock._buf.Length - 1;

            int b0 = prevBlock._buf[i++] & 0xFF;
            int b1 = _buf[_readIndex++] & 0xFF;
            return (b1 << 8) + (b0 << 0);
        }

        /**
         * Reads an <c>int</c> which was encoded in <em>little endian</em> format.
         */
        public int ReadIntLE()
        {
            int i = _readIndex;

            int b0 = _buf[i++] & 0xFF;
            int b1 = _buf[i++] & 0xFF;
            int b2 = _buf[i++] & 0xFF;
            int b3 = _buf[i++] & 0xFF;
            _readIndex = i;
            return (b3 << 24) + (b2 << 16) + (b1 << 8) + (b0 << 0);
        }

        /**
         * Reads an <c>int</c> which spans the end of <c>prevBlock</c> and the start of this block.
         */
        public int ReadIntLE(DataInputBlock prevBlock, int prevBlockAvailable)
        {
            byte[] buf = new byte[4];

            ReadSpanning(prevBlock, prevBlockAvailable, buf);
            int b0 = buf[0] & 0xFF;
            int b1 = buf[1] & 0xFF;
            int b2 = buf[2] & 0xFF;
            int b3 = buf[3] & 0xFF;
            return (b3 << 24) + (b2 << 16) + (b1 << 8) + (b0 << 0);
        }

        /**
         * Reads a <c>long</c> which was encoded in <em>little endian</em> format.
         */
        public long ReadLongLE()
        {
            int i = _readIndex;

            int b0 = _buf[i++] & 0xFF;
            int b1 = _buf[i++] & 0xFF;
            int b2 = _buf[i++] & 0xFF;
            int b3 = _buf[i++] & 0xFF;
            int b4 = _buf[i++] & 0xFF;
            int b5 = _buf[i++] & 0xFF;
            int b6 = _buf[i++] & 0xFF;
            int b7 = _buf[i++] & 0xFF;
            _readIndex = i;
            return (((long)b7 << 56) +
                    ((long)b6 << 48) +
                    ((long)b5 << 40) +
                    ((long)b4 << 32) +
                    ((long)b3 << 24) +
                    (b2 << 16) +
                    (b1 << 8) +
                    (b0 << 0));
        }

        /**
         * Reads a <c>long</c> which spans the end of <c>prevBlock</c> and the start of this block.
         */
        public long ReadLongLE(DataInputBlock prevBlock, int prevBlockAvailable)
        {
            byte[] buf = new byte[8];

            ReadSpanning(prevBlock, prevBlockAvailable, buf);

            int b0 = buf[0] & 0xFF;
            int b1 = buf[1] & 0xFF;
            int b2 = buf[2] & 0xFF;
            int b3 = buf[3] & 0xFF;
            int b4 = buf[4] & 0xFF;
            int b5 = buf[5] & 0xFF;
            int b6 = buf[6] & 0xFF;
            int b7 = buf[7] & 0xFF;
            return (((long)b7 << 56) +
                    ((long)b6 << 48) +
                    ((long)b5 << 40) +
                    ((long)b4 << 32) +
                    ((long)b3 << 24) +
                    (b2 << 16) +
                    (b1 << 8) +
                    (b0 << 0));
        }

        /**
         * Reads a small amount of data from across the boundary between two blocks.  
         * The {@link #_readIndex} of this (the second) block is updated accordingly.
         * Note- this method (and other code) assumes that the second {@link DataInputBlock}
         * always is big enough to complete the read without being exhausted.
         */
        private void ReadSpanning(DataInputBlock prevBlock, int prevBlockAvailable, byte[] buf)
        {
            Array.Copy(prevBlock._buf, prevBlock._readIndex, buf, 0, prevBlockAvailable);
            int secondReadLen = buf.Length - prevBlockAvailable;
            Array.Copy(_buf, 0, buf, prevBlockAvailable, secondReadLen);
            _readIndex = secondReadLen;
        }

        /**
         * Reads <c>len</c> bytes from this block into the supplied buffer.
         */
        public void ReadFully(byte[] buf, int off, int len)
        {
            Array.Copy(_buf, _readIndex, buf, off, len);
            _readIndex += len;
        }
    }
}

