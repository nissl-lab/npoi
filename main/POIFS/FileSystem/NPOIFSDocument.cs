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


using NPOI.POIFS.Common;
using NPOI.POIFS.Dev;
using NPOI.POIFS.Properties;
using NPOI.Util;
using System.IO;
using System.Collections.Generic;
using System;
using System.Text;
using System.Collections;

namespace NPOI.POIFS.FileSystem
{
    /**
 * This class manages a document in the NIO POIFS filesystem.
 * This is the {@link NPOIFSFileSystem} version.
 */
    public class NPOIFSDocument : POIFSViewable
    {
        private DocumentProperty _property;

        private NPOIFSFileSystem _filesystem;
        private NPOIFSStream _stream;
        private int _block_size;

        /**
         * Constructor for an existing Document 
         */
        public NPOIFSDocument(DocumentProperty property, NPOIFSFileSystem filesystem)
        {
            this._property = property;
            this._filesystem = filesystem;

            if (property.Size < POIFSConstants.BIG_BLOCK_MINIMUM_DOCUMENT_SIZE)
            {
                _stream = new NPOIFSStream(_filesystem.GetMiniStore(), property.StartBlock);
                _block_size = _filesystem.GetMiniStore().GetBlockStoreBlockSize();
            }
            else
            {
                _stream = new NPOIFSStream(_filesystem, property.StartBlock);
                _block_size = _filesystem.GetBlockStoreBlockSize();
            }
        }

        /**
         * Constructor for a new Document
         *
         * @param name the name of the POIFSDocument
         * @param stream the InputStream we read data from
         */
        public NPOIFSDocument(String name, NPOIFSFileSystem filesystem, Stream stream)
        {
            this._filesystem = filesystem;

            // Buffer the contents into memory. This is a bit icky...
            // TODO Replace with a buffer up to the mini stream size, then streaming write
            byte[] contents;
            if (stream is MemoryStream)
            {
                MemoryStream bais = (MemoryStream)stream;
                contents = new byte[bais.Length];
                bais.Read(contents, 0, contents.Length);
            }
            else
            {
                MemoryStream baos = new MemoryStream();
                IOUtils.Copy(stream, baos);
                contents = baos.ToArray();
            }

            // Do we need to store as a mini stream or a full one?
            if (contents.Length <= POIFSConstants.BIG_BLOCK_MINIMUM_DOCUMENT_SIZE)
            {
                _stream = new NPOIFSStream(filesystem.GetMiniStore());
                _block_size = _filesystem.GetMiniStore().GetBlockStoreBlockSize();
            }
            else
            {
                _stream = new NPOIFSStream(filesystem);
                _block_size = _filesystem.GetBlockStoreBlockSize();
            }

            // Store it
            _stream.UpdateContents(contents);

            // And build the property for it
            this._property = new DocumentProperty(name, contents.Length);
            _property.StartBlock = _stream.GetStartBlock();
        }

        public int GetDocumentBlockSize()
        {
            return _block_size;
        }

        public IEnumerator<ByteBuffer> GetBlockIterator()
        {
            if (Size > 0)
            {
                return _stream.GetBlockIterator();
            }
            else
            {
                //List<byte[]> empty = Collections.emptyList();
                List<ByteBuffer> empty = new List<ByteBuffer>();
                return empty.GetEnumerator();
            }
        }

        /**
         * @return size of the document
         */
        public int Size
        {
            get
            {
                return _property.Size;
            }
        }

        /**
         * @return the instance's DocumentProperty
         */
        public DocumentProperty DocumentProperty
        {
            get
            {
                return _property;
            }
        }

        /**
         * Get an array of objects, some of which may implement POIFSViewable
         *
         * @return an array of Object; may not be null, but may be empty
         */
        protected Object[] GetViewableArray()
        {
            Object[] results = new Object[1];
            String result;

            try
            {
                if (Size > 0)
                {
                    // Get all the data into a single array
                    byte[] data = new byte[Size];
                    int offset = 0;
                    foreach (ByteBuffer buffer in _stream)
                    {
                        int length = Math.Min(_block_size, data.Length - offset);
                        buffer.Read(data, offset, length);
                        offset += length;
                    }

                    MemoryStream output = new MemoryStream();
                    HexDump.Dump(data, 0, output, 0);
                    result = output.ToString();
                }
                else
                {
                    result = "<NO DATA>";
                }
            }
            catch (IOException e)
            {
                result = e.Message;
            }
            results[0] = result;
            return results;
        }

        /**
              * Get an Iterator of objects, some of which may implement POIFSViewable
              *
              * @return an Iterator; may not be null, but may have an empty back end
              *		 store
              */
        protected IEnumerator GetViewableIterator()
        {
            //  return Collections.EMPTY_LIST.iterator();
            return null;

        }


        /**
    * Provides a short description of the object, to be used when a
    * POIFSViewable object has not provided its contents.
    *
    * @return short description
    */
        protected String GetShortDescription()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("Document: \"").Append(_property.Name).Append("\"");
            buffer.Append(" size = ").Append(Size);
            return buffer.ToString();
        }

        #region POIFSViewable Members

        public bool PreferArray
        {
            get { return true; }
        }

        public string ShortDescription
        {
            get { return GetShortDescription(); }
        }

        public Array ViewableArray
        {
            get { return GetViewableArray(); }
        }

        public IEnumerator ViewableIterator
        {
            get { return GetViewableIterator(); }
        }

        #endregion
    }
}