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
    using System.IO;
    using NPOI.Util;

    /**
     * This class provides methods to read a DocumentEntry managed by a
     *  {@link POIFSFileSystem} or {@link NPOIFSFileSystem} instance.
     * It Creates the appropriate one, and delegates, allowing us to
     *  work transparently with the two.
     */
    public class DocumentInputStream : ByteArrayInputStream, ILittleEndianInput
    {
        /** returned by read operations if we're at end of document */
        protected static int EOF = -1;

        protected static int SIZE_SHORT = 2;
        protected static int SIZE_INT = 4;
        protected static int SIZE_LONG = 8;

        private DocumentInputStream delegate1;

        /** For use by downstream implementations */
        protected DocumentInputStream() { }

        /**
         * Create an InputStream from the specified DocumentEntry
         * 
         * @param document the DocumentEntry to be read
         * 
         * @exception IOException if the DocumentEntry cannot be opened (like, maybe it has
         *                been deleted?)
         */
        public DocumentInputStream(DocumentEntry document)
        {
            if (!(document is DocumentNode))
            {
                throw new IOException("Cannot open internal document storage");
            }
            DocumentNode documentNode = (DocumentNode)document;
            DirectoryNode parentNode = (DirectoryNode)document.Parent;

            if (documentNode.Document != null)
            {
                delegate1 = new ODocumentInputStream(document);
            }
            else if (parentNode.FileSystem != null)
            {
                delegate1 = new ODocumentInputStream(document);
            }
            else if (parentNode.NFileSystem != null)
            {
                delegate1 = new NDocumentInputStream(document);
            }
            else
            {
                throw new IOException("No FileSystem bound on the parent, can't read contents");
            }
        }
        public override long Seek(long offset, SeekOrigin origin)
        {
            return delegate1.Seek(offset, origin);
        }

        public override long Length
        {
            get
            {
                return delegate1.Length;
            }
        }
        public override long Position
        {
            get
            {
                return delegate1.Position;
            }
            set
            {
                delegate1.Position = value;
            }
        }
        /**
         * Create an InputStream from the specified Document
         * 
         * @param document the Document to be read
         */
        public DocumentInputStream(POIFSDocument document)
        {
            delegate1 = new ODocumentInputStream(document);
        }

        /**
         * Create an InputStream from the specified Document
         * 
         * @param document the Document to be read
         */
        public DocumentInputStream(NPOIFSDocument document)
        {
            delegate1 = new NDocumentInputStream(document);
        }

        public override int Available()
        {
            return delegate1.Available();
        }

        public override void Close()
        {
            delegate1.Close();
        }

        public override void Mark(int ignoredReadlimit)
        {
            delegate1.Mark(ignoredReadlimit);
        }

        /**
         * Tests if this input stream supports the mark and reset methods.
         * 
         * @return <code>true</code> always
         */
        public override bool MarkSupported()
        {
            return true;
        }

        public override int Read()
        {
            return delegate1.Read();
        }

        public virtual int Read(byte[] b)
        {
            return Read(b, 0, b.Length);
        }

        public override int Read(byte[] b, int off, int len)
        {
            return delegate1.Read(b, off, len);
        }

        /**
         * Repositions this stream to the position at the time the mark() method was
         * last called on this input stream. If mark() has not been called this
         * method repositions the stream to its beginning.
         */
        public override void Reset()
        {
            delegate1.Reset();
        }

        public virtual long Skip(long n)
        {
            return delegate1.Skip(n);
        }

        public override int ReadByte()
        {
            return delegate1.ReadByte();
        }

        public virtual double ReadDouble()
        {
            return delegate1.ReadDouble();
        }

        public virtual short ReadShort()
        {
            return (short)ReadUShort();
        }

        public virtual void ReadFully(byte[] buf)
        {
            ReadFully(buf, 0, buf.Length);
        }

        public virtual void ReadFully(byte[] buf, int off, int len)
        {
            delegate1.ReadFully(buf, off, len);
        }

        public virtual long ReadLong()
        {
            return delegate1.ReadLong();
        }

        public virtual int ReadInt()
        {
            return delegate1.ReadInt();
        }

        public virtual int ReadUShort()
        {
            return delegate1.ReadUShort();
        }

        public virtual int ReadUByte()
        {
            return delegate1.ReadUByte();
        }
    }

}