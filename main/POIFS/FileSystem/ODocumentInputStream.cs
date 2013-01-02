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
using NPOI.POIFS.Storage;
using System.IO;

namespace NPOI.POIFS.FileSystem
{
    /**
     * This class provides methods to read a DocumentEntry managed by a
     * {@link POIFSFileSystem} instance.
     *
     * @author Marc Johnson (mjohnson at apache dot org)
     */
    public class ODocumentInputStream : DocumentInputStream//DocumentReader
    {
        /** current offset into the Document */
        private long _current_offset;

        /** current marked offset into the Document (used by mark and Reset) */
        private long _marked_offset;

        /** the Document's size */
        private int _document_size;

        /** have we been closed? */
        private bool _closed;

        /** the actual Document */
        private POIFSDocument _document;

        /** the data block Containing the current stream pointer */
        private DataInputBlock _currentBlock;

        /**
         * Create an InputStream from the specified DocumentEntry
         * 
         * @param document the DocumentEntry to be read
         * 
         * @exception IOException if the DocumentEntry cannot be opened (like, maybe it has
         *                been deleted?)
         */
        public ODocumentInputStream(DocumentEntry document)
        {
            if (!(document is DocumentNode))
            {
                throw new IOException("Cannot open internal document storage");
            }
            DocumentNode documentNode = (DocumentNode)document;
            if (documentNode.Document == null)
            {
                throw new IOException("Cannot open internal document storage");
            }

            _current_offset = 0;
            _marked_offset = 0;
            _document_size = document.Size;
            _closed = false;
            _document = documentNode.Document;
            _currentBlock = GetDataInputBlock(0);
        }

        public override long Length
        {
            get
            {
                return _document_size;
            }
        }
        
        /**
         * Create an InputStream from the specified Document
         * 
         * @param document the Document to be read
         */
        public ODocumentInputStream(POIFSDocument document)
        {
            _current_offset = 0;
            _marked_offset = 0;
            _document_size = document.Size;
            _closed = false;
            _document = document;
            _currentBlock = GetDataInputBlock(0);
        }


        public override int Available()
        {
            if (_closed)
            {
                throw new InvalidOperationException("cannot perform requested operation on a closed stream");
            }
            return _document_size - (int)_current_offset;
        }


        public override void Close()
        {
            _closed = true;
        }


        public override void Mark(int ignoredReadlimit)
        {
            _marked_offset = _current_offset;
        }

        private DataInputBlock GetDataInputBlock(long offset)
        {
            return _document.GetDataInputBlock((int)offset);
        }


        public override int Read()
        {
            dieIfClosed();
            if (atEOD())
            {
                return EOF;
            }
            int result = _currentBlock.ReadUByte();
            _current_offset++;
            if (_currentBlock.Available() < 1)
            {
                _currentBlock = GetDataInputBlock(_current_offset);
            }
            return result;
        }


        public override int Read(byte[] b, int off, int len)
        {
            dieIfClosed();
            if (b == null)
            {
                throw new ArgumentException("buffer must not be null");
            }
            if (off < 0 || len < 0 || b.Length < off + len)
            {
                throw new IndexOutOfRangeException("can't read past buffer boundaries");
            }
            if (len == 0)
            {
                return 0;
            }
            if (atEOD())
            {
                return EOF;
            }
            int limit = Math.Min(Available(), len);
            ReadFully(b, off, limit);
            return limit;
        }

        /**
         * Repositions this stream to the position at the time the mark() method was
         * last called on this input stream. If mark() has not been called this
         * method repositions the stream to its beginning.
         */

        public override void Reset()
        {
            _current_offset = _marked_offset;
            _currentBlock = GetDataInputBlock(_current_offset);
        }


        public override long Skip(long n)
        {
            dieIfClosed();
            if (n < 0)
            {
                return 0;
            }
            long new_offset = _current_offset + (int)n;

            if (new_offset < _current_offset)
            {

                // wrap around in Converting a VERY large long to an int
                new_offset = _document_size;
            }
            else if (new_offset > _document_size)
            {
                new_offset = _document_size;
            }
            long rval = new_offset - _current_offset;

            _current_offset = new_offset;
            _currentBlock = GetDataInputBlock(_current_offset);
            return rval;
        }

        private void dieIfClosed()
        {
            if (_closed)
            {
                throw new IOException("cannot perform requested operation on a closed stream");
            }
        }

        private bool atEOD()
        {
            return _current_offset == _document_size;
        }

        private void CheckAvaliable(int requestedSize)
        {
            if (_closed)
            {
                throw new InvalidOperationException("cannot perform requested operation on a closed stream");
            }
            if (requestedSize > _document_size - _current_offset)
            {
                throw new Exception("Buffer underrun - requested " + requestedSize
                        + " bytes but " + (_document_size - _current_offset) + " was available");
            }
        }


        public override int ReadByte()
        {
            return ReadUByte();
        }


        public override double ReadDouble()
        {
            return BitConverter.Int64BitsToDouble(ReadLong());
        }


        public override short ReadShort()
        {
            return (short)ReadUShort();
        }


        public override void ReadFully(byte[] buf, int off, int len)
        {
            CheckAvaliable(len);
            int blockAvailable = _currentBlock.Available();
            if (blockAvailable > len)
            {
                _currentBlock.ReadFully(buf, off, len);
                _current_offset += len;
                return;
            }
            // else read big amount in chunks
            int remaining = len;
            int WritePos = off;
            while (remaining > 0)
            {
                bool blockIsExpiring = remaining >= blockAvailable;
                int reqSize;
                if (blockIsExpiring)
                {
                    reqSize = blockAvailable;
                }
                else
                {
                    reqSize = remaining;
                }
                _currentBlock.ReadFully(buf, WritePos, reqSize);
                remaining -= reqSize;
                WritePos += reqSize;
                _current_offset += reqSize;
                if (blockIsExpiring)
                {
                    if (_current_offset == _document_size)
                    {
                        if (remaining > 0)
                        {
                            throw new InvalidOperationException(
                                    "reached end of document stream unexpectedly");
                        }
                        _currentBlock = null;
                        break;
                    }
                    _currentBlock = GetDataInputBlock(_current_offset);
                    blockAvailable = _currentBlock.Available();
                }
            }
        }


        public override long ReadLong()
        {
            CheckAvaliable(SIZE_LONG);
            int blockAvailable = _currentBlock.Available();
            long result;
            if (blockAvailable > SIZE_LONG)
            {
                result = _currentBlock.ReadLongLE();
            }
            else
            {
                DataInputBlock nextBlock = GetDataInputBlock(_current_offset + blockAvailable);
                if (blockAvailable == SIZE_LONG)
                {
                    result = _currentBlock.ReadLongLE();
                }
                else
                {
                    result = nextBlock.ReadLongLE(_currentBlock, blockAvailable);
                }
                _currentBlock = nextBlock;
            }
            _current_offset += SIZE_LONG;
            return result;
        }


        public override int ReadInt()
        {
            CheckAvaliable(SIZE_INT);
            int blockAvailable = _currentBlock.Available();
            int result;
            if (blockAvailable > SIZE_INT)
            {
                result = _currentBlock.ReadIntLE();
            }
            else
            {
                DataInputBlock nextBlock = GetDataInputBlock(_current_offset + blockAvailable);
                if (blockAvailable == SIZE_INT)
                {
                    result = _currentBlock.ReadIntLE();
                }
                else
                {
                    result = nextBlock.ReadIntLE(_currentBlock, blockAvailable);
                }
                _currentBlock = nextBlock;
            }
            _current_offset += SIZE_INT;
            return result;
        }


        public override int ReadUShort()
        {
            CheckAvaliable(SIZE_SHORT);
            int blockAvailable = _currentBlock.Available();
            int result;
            if (blockAvailable > SIZE_SHORT)
            {
                result = _currentBlock.ReadUshortLE();
            }
            else
            {
                DataInputBlock nextBlock = GetDataInputBlock(_current_offset + blockAvailable);
                if (blockAvailable == SIZE_SHORT)
                {
                    result = _currentBlock.ReadUshortLE();
                }
                else
                {
                    result = nextBlock.ReadUshortLE(_currentBlock);
                }
                _currentBlock = nextBlock;
            }
            _current_offset += SIZE_SHORT;
            return result;
        }


        public override int ReadUByte()
        {
            CheckAvaliable(1);
            int result = _currentBlock.ReadUByte();
            _current_offset++;
            if (_currentBlock.Available() < 1)
            {
                _currentBlock = GetDataInputBlock(_current_offset);
            }
            return result;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            if (origin == SeekOrigin.Current)
            {
                if (_current_offset + offset >= this.Length || _current_offset + offset < 0)
                    throw new ArgumentException("invalid offset");
                _current_offset += (int)offset;
            }
            else if (origin == SeekOrigin.Begin)
            {
                if (offset >= this.Length || offset < 0)
                    throw new ArgumentException("invalid offset");

                _current_offset = offset;
            }
            else if (origin == SeekOrigin.End)
            {
                if (this.Length + offset >= this.Length || this.Length + offset < 0)
                    throw new ArgumentException("invalid offset");

                _current_offset = this.Length + offset;
            }
            return _current_offset;
        }

        public override long Position
        {
            get
            {
                if (_closed)
                {
                    throw new InvalidOperationException("cannot perform requested operation on a closed stream");
                }
                return _current_offset;
            }
            set
            {
                _current_offset = (int)value;
            }
        }
    }

}





