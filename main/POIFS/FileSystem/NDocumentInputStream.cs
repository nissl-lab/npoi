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



using NPOI.POIFS.Properties;
using NPOI.Util;
using System.IO;
using System;
using System.Collections.Generic;


namespace NPOI.POIFS.FileSystem
{


    /**
     * This class provides methods to read a DocumentEntry managed by a
     * {@link NPOIFSFileSystem} instance.
     */
    public class NDocumentInputStream : DocumentInputStream//DocumentReader
    {
        /** current offset into the Document */
        private int _current_offset;
        /** current block count */
        private int _current_block_count;

        /** current marked offset into the Document (used by mark and Reset) */
        private int _marked_offset;
        /** and the block count for it */
        private int _marked_offset_count;

        /** the Document's size */
        private int _document_size;

        /** have we been closed? */
        private bool _closed;

        /** the actual Document */
        private NPOIFSDocument _document;

        private IEnumerator<ByteBuffer> _data;
        private ByteBuffer _buffer;

        

        /**
         * Create an InputStream from the specified DocumentEntry
         * 
         * @param document the DocumentEntry to be read
         * 
         * @exception IOException if the DocumentEntry cannot be opened (like, maybe it has
         *                been deleted?)
         */
        public NDocumentInputStream(DocumentEntry document)
        {
            if (!(document is DocumentNode))
            {
                throw new IOException("Cannot open internal document storage, " + document + " not a Document Node");
            }
            _current_offset = 0;
            _current_block_count = 0;
            _marked_offset = 0;
            _marked_offset_count = 0;
            _document_size = document.Size;
            _closed = false;

            DocumentNode doc = (DocumentNode)document;
            DocumentProperty property = (DocumentProperty)doc.Property;
            _document = new NPOIFSDocument(
                  property,
                      ((DirectoryNode)doc.Parent).NFileSystem
            );
            _data = _document.GetBlockIterator();
        }

        /**
         * Create an InputStream from the specified Document
         * 
         * @param document the Document to be read
         */
        public NDocumentInputStream(NPOIFSDocument document)
        {
            _current_offset = 0;
            _current_block_count = 0;
            _marked_offset = 0;
            _marked_offset_count = 0;
            _document_size = document.Size;
            _closed = false;
            _document = document;
            _data = _document.GetBlockIterator();
        }


        public override int Available()
        {
            if (_closed)
            {
                throw new InvalidOperationException("cannot perform requested operation on a closed stream");
            }
            return _document_size - _current_offset;
        }


        public override void Close()
        {
            _closed = true;
        }


        public override void Mark(int ignoredReadlimit)
        {
            _marked_offset = _current_offset;
            _marked_offset_count = Math.Max(0, _current_block_count - 1);
        }


        public override int Read()
        {
            DieIfClosed();
            if (atEOD())
            {
                return EOF;
            }
            byte[] b = new byte[1];
            int result = Read(b, 0, 1);
            if (result >= 0)
            {
                if (b[0] < 0)
                {
                    return b[0] + 256;
                }
                return b[0];
            }
            return result;
        }


        public override int Read(byte[] b, int off, int len)
        {
            DieIfClosed();
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
            // Special case for Reset to the start
            if (_marked_offset == 0 && _marked_offset_count == 0)
            {
                _current_block_count = _marked_offset_count;
                _current_offset = _marked_offset;
                _data = _document.GetBlockIterator();
                _buffer = null;
                return;
            }

            // Start again, then wind on to the required block
            _data = _document.GetBlockIterator();
            _current_offset = 0;

            for (int i = 0; i < _marked_offset_count; i++)
            {
                _data.MoveNext();
                _buffer = _data.Current;
                _current_offset += _buffer.Remain;
            }

            _current_block_count = _marked_offset_count;

            // Do we need to position within it?
            if (_current_offset != _marked_offset)
            {
                // Grab the right block
                _data.MoveNext();
                _buffer = _data.Current;
                _current_block_count++;

                // Skip to the right place in it
                // (It should be positioned already at the start of the block,
                //  we need to Move further inside the block)
                int skipBy = _marked_offset - _current_offset;
                _buffer.Position = (_buffer.Position + skipBy);
            }

            // All done
            _current_offset = _marked_offset;
        }


        public override long Skip(long n)
        {
            DieIfClosed();
            if (n < 0)
            {
                return 0;
            }
            int new_offset = _current_offset + (int)n;

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

            // TODO Do this better
            byte[] Skip = new byte[(int)rval];
            ReadFully(Skip);
            return rval;
        }

        private void DieIfClosed()
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


        public override void ReadFully(byte[] buf, int off, int len)
        {
            CheckAvaliable(len);

            int read = 0;
            while (read < len)
            {
                // if (_buffer == null || _buffer.remaining() == 0)
                if (_buffer == null || _buffer.Remain == 0)
                {
                    _current_block_count++;
                    //_buffer = _data.next();
                    _data.MoveNext();
                    _buffer = _data.Current;

                }

                int limit = Math.Min(len - read, _buffer.Remain);
                _buffer.Read(buf, off + read, limit);
                _current_offset += limit;
                read += limit;
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


        public override long ReadLong()
        {
            CheckAvaliable(SIZE_LONG);
            byte[] data = new byte[SIZE_LONG];
            ReadFully(data, 0, SIZE_LONG);
            return LittleEndian.GetLong(data, 0);
        }

        public override void ReadFully(byte[] buf)
        {
            ReadFully(buf, 0, buf.Length);
        }


        public override short ReadShort()
        {
            CheckAvaliable(SIZE_SHORT);
            byte[] data = new byte[SIZE_SHORT];
            ReadFully(data, 0, SIZE_SHORT);
            return LittleEndian.GetShort(data);
        }


        public override int ReadInt()
        {
            CheckAvaliable(SIZE_INT);
            byte[] data = new byte[SIZE_INT];
            ReadFully(data, 0, SIZE_INT);
            return LittleEndian.GetInt(data);
        }


        public override int ReadUShort()
        {
            CheckAvaliable(SIZE_SHORT);
            byte[] data = new byte[SIZE_SHORT];
            ReadFully(data, 0, SIZE_SHORT);
            return LittleEndian.GetUShort(data);
        }


        public override int ReadUByte()
        {
            CheckAvaliable(1);
            byte[] data = new byte[1];
            ReadFully(data, 0, 1);
            if (data[0] >= 0)
                return data[0];
            return data[0] + 256;
        }

        public override long Length
        {
            get
            {
                if (_closed)
                {
                    throw new InvalidOperationException("cannot perform requested operation on a closed stream");
                }
                return _document_size;   //- _current_offset;
            }
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

        public override long Seek(long offset, SeekOrigin origin)
        {
            if (offset == 0)
            {
                Reset();

            }
            else
            {
                Mark((int)offset);
            }

            return 0;
        }
    }
}