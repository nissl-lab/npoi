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

namespace NPOI.POIFS.FileSystem
{
    using System;
    using System.IO;
    using NPOI.POIFS.Properties;
    using NPOI.POIFS.Common;

    /**
     * This class provides methods to write a DocumentEntry managed by a
     * {@link NPOIFSFileSystem} instance.
     */
    public class NDocumentOutputStream : MemoryStream
    {
        /** the Document's size */
        private int _document_size;

        /** have we been closed? */
        private bool _closed;

        /** the actual Document */
        private NPOIFSDocument _document;

        /** and its Property */
        private DocumentProperty _property;

        /** our buffer, when null we're into normal blocks */
        private MemoryStream _buffer =
                new MemoryStream(POIFSConstants.BIG_BLOCK_MINIMUM_DOCUMENT_SIZE);

        /** our main block stream, when we're into normal blocks */
        private NPOIFSStream _stream;
        private Stream _stream_output;
        /**
         * Create an OutputStream from the specified DocumentEntry.
         * The specified entry will be emptied.
         * 
         * @param document the DocumentEntry to be written
         */
        public NDocumentOutputStream(DocumentEntry document)
        {
            if (!(document is DocumentNode))
            {
                throw new IOException("Cannot open internal document storage, " + document + " not a Document Node");
            }
            _document_size = 0;
            _closed = false;

            _property = (DocumentProperty)((DocumentNode)document).Property;

            _document = new NPOIFSDocument((DocumentNode)document);
            _document.Free();
        }

        /**
         * Create an OutputStream to create the specified new Entry
         * 
         * @param parent Where to create the Entry
         * @param name Name of the new entry
         */
        public NDocumentOutputStream(DirectoryEntry parent, String name)
        {
            if (!(parent is DirectoryNode))
            {
                throw new IOException("Cannot open internal directory storage, " + parent + " not a Directory Node");
            }
            _document_size = 0;
            _closed = false;

            // Have an empty one Created for now
            DocumentEntry doc = parent.CreateDocument(name, new MemoryStream(new byte[0]));
            _property = (DocumentProperty)((DocumentNode)doc).Property;
            _document = new NPOIFSDocument((DocumentNode)doc);
        }
        private void dieIfClosed()
        {
            if (_closed)
            {
                throw new IOException("cannot perform requested operation on a closed stream");
            }
        }

        private void CheckBufferSize()
        {
            // Have we gone over the mini stream limit yet?
            if (_buffer.Length > POIFSConstants.BIG_BLOCK_MINIMUM_DOCUMENT_SIZE)
            {
                // Will need to be in the main stream
                byte[] data = _buffer.ToArray();
                _buffer = null;
                Write(data, 0, data.Length);
            }
            else
            {
                // So far, mini stream will work, keep going
            }
        }
        public void Write(int b)
        {
            dieIfClosed();

            if (_buffer != null)
            {
                _buffer.WriteByte((byte)b);
                CheckBufferSize();
            }
            else
            {
                Write(new byte[] { (byte)b });
            }
        }

        public void Write(byte[] b)
        {
            dieIfClosed();

            if (_buffer != null)
            {
                _buffer.Write(b, 0, b.Length);
                CheckBufferSize();
            }
            else
            {
                Write(b, 0, b.Length);
            }
        }

        public override void Write(byte[] b, int off, int len)
        {
            dieIfClosed();

            if (_buffer != null)
            {
                _buffer.Write(b, off, len);
                CheckBufferSize();
            }
            else
            {
                if (_stream == null)
                {
                    _stream = new NPOIFSStream(_document.FileSystem);
                    _stream_output = _stream.GetOutputStream();
                }
                _stream_output.Write(b, off, len);
                _document_size += len;
            }
        }

        public override void Close()
        {
            base.Close();
            // Do we have a pending buffer for the mini stream?
            if (_buffer != null)
            {
                // It's not much data, so ask NPOIFSDocument to do it for us
                _document.ReplaceContents(new MemoryStream(_buffer.ToArray()));
            }
            else
            {
                // We've been writing to the stream as we've gone along
                // Update the details on the property now
                _stream_output.Close();
                _property.UpdateSize(_document_size);
                _property.StartBlock=(_stream.GetStartBlock());
            }

            // No more!
            _closed = true;
        }
    }

}