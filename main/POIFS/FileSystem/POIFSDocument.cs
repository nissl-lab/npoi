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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/


using System;
using System.Text;
using System.IO;
using System.Collections;

using NPOI.POIFS.Common;
using NPOI.POIFS.Storage;
using NPOI.POIFS.Dev;
using NPOI.POIFS.Properties;
using NPOI.POIFS.EventFileSystem;
using NPOI.Util;

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// This class manages a document in the POIFS filesystem.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class POIFSDocument : BATManaged, BlockWritable, POIFSViewable,IDisposable
    {
        private BigBlockStore _big_store;
        private DocumentProperty _property;
        private int _size;
        private SmallBlockStore _small_store;

        /// <summary>
        /// Initializes a new instance of the <see cref="POIFSDocument"/> class.
        /// </summary>
        /// <param name="name">the name of the POIFSDocument</param>
        /// <param name="stream">the InputStream we read data from</param>
        public POIFSDocument(string name, Stream stream)
        {
            DocumentBlock block;
            IList list = new ArrayList();
            this._size = 0;
            do
            {
                block = new DocumentBlock(stream);
                int size = block.Size;
                if (size > 0)
                {
                    list.Add(block);
                    this._size += size;
                }
            }
            while (!block.PartiallyRead);
            DocumentBlock[] blocks = (DocumentBlock[])((ArrayList)list).ToArray(typeof(DocumentBlock));
            this._big_store = new BigBlockStore(this, blocks);
            this._property = new DocumentProperty(name, this._size);
            this._property.Document = this;
            if (this._property.ShouldUseSmallBlocks)
            {
                this._small_store = new SmallBlockStore(this, SmallDocumentBlock.Convert(blocks, this._size));
                this._big_store = new BigBlockStore(this, new DocumentBlock[0]);
            }
            else
            {
                this._small_store = new SmallBlockStore(this, new BlockWritable[0]);
            }
        }
        public void Dispose()
        {
            _big_store = null;
            _property = null;
            _small_store = null;
        }
        /// <summary>
        /// Constructor from small blocks
        /// </summary>
        /// <param name="name">the name of the POIFSDocument</param>
        /// <param name="blocks">the small blocks making up the POIFSDocument</param>
        /// <param name="length">the actual length of the POIFSDocument</param>
        public POIFSDocument(string name, ListManagedBlock[] blocks, int length)
        {
            this._size = length;
            this._property = new DocumentProperty(name, this._size);
            this._property.Document = this;
            if (Property.IsSmall(this._size))
            {
                this._big_store = new BigBlockStore(this, new RawDataBlock[0]);
                this._small_store = new SmallBlockStore(this, blocks);
            }
            else
            {
                this._big_store = new BigBlockStore(this, blocks);
                this._small_store = new SmallBlockStore(this, new BlockWritable[0]);
            }
        }
        /// <summary>
        /// Constructor from large blocks
        /// </summary>
        /// <param name="name">the name of the POIFSDocument</param>
        /// <param name="blocks">the big blocks making up the POIFSDocument</param>
        /// <param name="length">the actual length of the POIFSDocument</param>
        public POIFSDocument(string name, RawDataBlock[] blocks, int length)
        {
            this._size = length;
            this._big_store = new BigBlockStore(this, blocks);
            this._property = new DocumentProperty(name, this._size);
            this._small_store = new SmallBlockStore(this, new BlockWritable[0]);
            this._property.Document = this;
        }
        /// <summary>
        /// Constructor from small blocks
        /// </summary>
        /// <param name="name">the name of the POIFSDocument</param>
        /// <param name="blocks">Tthe small blocks making up the POIFSDocument</param>
        /// <param name="length">the actual length of the POIFSDocument</param>
        public POIFSDocument(string name, SmallDocumentBlock[] blocks, int length)
        {
            this._size = length;
            try
            {
                this._big_store = new BigBlockStore(this, new RawDataBlock[0]);
            }
            catch (IOException)
            {
            }
            this._property = new DocumentProperty(name, this._size);
            this._small_store = new SmallBlockStore(this, blocks);
            this._property.Document = this;
        }

        /// <summary>
        /// Occurs when [before writing].
        /// </summary>
        public event POIFSWriterEventHandler BeforeWriting
        {
            add {
                if (_property.ShouldUseSmallBlocks)
                {
                    _small_store.BeforeWriting += value;
                }
                else
                {
                    _big_store.BeforeWriting += value;
                }
            }
            remove
            {
                if (_property.ShouldUseSmallBlocks)
                {
                    _small_store.BeforeWriting -= value;
                }
                else
                {
                    _big_store.BeforeWriting -= value;
                }
            }
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="POIFSDocument"/> class.
        /// </summary>
        /// <param name="name">the name of the POIFSDocument</param>
        /// <param name="size">the length of the POIFSDocument</param>
        /// <param name="path">the path of the POIFSDocument</param>
        public POIFSDocument(string name, int size, POIFSDocumentPath path)
        {
            this._size = size;
            this._property = new DocumentProperty(name, this._size);
            this._property.Document = this;
            if (this._property.ShouldUseSmallBlocks)
            {
                this._small_store = new SmallBlockStore(this, path, name, size);
                this._big_store = new BigBlockStore(this, new object[0]);
            }
            else
            {
                this._small_store = new SmallBlockStore(this, new BlockWritable[0]);
                this._big_store = new BigBlockStore(this, path, name, size);
            }
        }

        /// <summary>
        /// read data from the internal stores
        /// </summary>
        /// <param name="buffer">the buffer to write to</param>
        /// <param name="offset">the offset into our storage to read from</param>
        public virtual void Read(byte[] buffer, int offset)
        {
            if (this._property.ShouldUseSmallBlocks)
            {
                SmallDocumentBlock.Read(this._small_store.Blocks, buffer, offset);
            }
            else
            {
                DocumentBlock.Read(this._big_store.Blocks, buffer, offset);
            }
        }

        /// <summary>
        /// Writes the blocks.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public virtual void WriteBlocks(Stream stream)
        {
            this._big_store.WriteBlocks(stream);
        }


        /// <summary>
        /// Gets the number of BigBlock's this instance uses
        /// </summary>
        /// <value>count of BigBlock instances</value>
        public virtual int CountBlocks
        {
            get
            {
                return this._big_store.CountBlocks;
            }
        }

        /// <summary>
        /// Gets the document property.
        /// </summary>
        /// <value>The document property.</value>
        public virtual DocumentProperty DocumentProperty
        {
            get
            {
                return this._property;
            }
        }

        /// <summary>
        /// Provides a short description of the object to be used when a
        /// POIFSViewable object has not provided its contents.
        /// </summary>
        /// <value><c>true</c> if [prefer array]; otherwise, <c>false</c>.</value>
        public virtual bool PreferArray
        {
            get
            {
                return true;
            }
        }

        /// <summary>
        /// Gets the short description.
        /// </summary>
        /// <value>The short description.</value>
        public virtual string ShortDescription
        {
            get
            {
                StringBuilder builder = new StringBuilder();
                builder.Append("Document: \"").Append(this._property.Name).Append("\"");
                builder.Append(" size = ").Append(this.Size);
                return builder.ToString();
            }
        }

        /// <summary>
        /// Gets the size.
        /// </summary>
        /// <value>The size.</value>
        public virtual int Size
        {
            get
            {
                return this._size;
            }
        }

        /// <summary>
        /// Gets the small blocks.
        /// </summary>
        /// <value>The small blocks.</value>
        public virtual BlockWritable[] SmallBlocks
        {
            get
            {
                return this._small_store.Blocks;
            }
        }

        /// <summary>
        /// Sets the start block for this instance
        /// </summary>
        /// <value>
        /// index into the array of BigBlock instances making up the the filesystem
        /// </value>
        public virtual int StartBlock
        {
            get
            {
                return this._property.StartBlock;
            }
            set
            {
                this._property.StartBlock = value;
            }
        }

        /// <summary>
        /// Get an array of objects, some of which may implement POIFSViewable
        /// </summary>
        /// <value>The viewable array.</value>
        public Array ViewableArray
        {
            get
            {
                string message;
                object[] objArray = new object[1];
                try
                {
                    MemoryStream stream = new MemoryStream();
                    BlockWritable[] blocks = null;
                    if (this._big_store.Valid)
                    {
                        blocks = this._big_store.Blocks;
                    }
                    else if (this._small_store.Valid)
                    {
                        blocks = this._small_store.Blocks;
                    }
                    if (blocks != null)
                    {
                        for (int i = 0; i < blocks.Length; i++)
                        {
                            blocks[i].WriteBlocks(stream);
                        }
                        byte[] sourceArray = stream.ToArray();
                        if (sourceArray.Length > this._property.Size)
                        {
                            byte[] buffer2 = new byte[this._property.Size];
                            Array.Copy(sourceArray, 0, buffer2, 0, buffer2.Length);
                            sourceArray = buffer2;
                        }
                        stream = new MemoryStream();
                        HexDump.Dump(sourceArray, 0L, stream, 0);
                        byte[] buffer = stream.GetBuffer();
                        char[] destinationArray = new char[(int)stream.Length];
                        Array.Copy(buffer, 0, destinationArray, 0, destinationArray.Length);
                        message = new string(destinationArray);
                    }
                    else
                    {
                        message = "<NO DATA>";
                    }
                }
                catch (IOException exception)
                {
                    message = exception.Message;
                }
                objArray[0] = message;
                return objArray;
            }
        }

        /// <summary>
        /// Give viewers a hint as to whether to call ViewableArray or ViewableIterator
        /// </summary>
        /// <value>The viewable iterator.</value>
        public virtual IEnumerator ViewableIterator
        {
            get
            {
                return ArrayList.ReadOnly(new ArrayList()).GetEnumerator();
            }
        }

        // Nested Types
        internal class BigBlockStore:IDisposable
        {
            private DocumentBlock[] bigBlocks;
            private POIFSDocument enclosingInstance;
            private string name;
            private POIFSDocumentPath path;
            private int size;

            public void Dispose()
            {
                bigBlocks = null;
                enclosingInstance = null;
                name = null;
                path = null;
            }

            internal BigBlockStore(POIFSDocument enclosingInstance, object[] blocks)
            {
                this.InitBlock(enclosingInstance);
                this.bigBlocks = new DocumentBlock[blocks.Length];
                for (int i = 0; i < blocks.Length; i++)
                {
                    if (blocks[i] is DocumentBlock)
                    {
                        this.bigBlocks[i] = (DocumentBlock)blocks[i];
                    }
                    else
                    {
                        this.bigBlocks[i] = new DocumentBlock((RawDataBlock)blocks[i]);
                    }
                }
                this.path = null;
                this.name = null;
                this.size = -1;
            }

            internal BigBlockStore(POIFSDocument enclosingInstance, POIFSDocumentPath path, string name, int size)
            {
                this.InitBlock(enclosingInstance);
                this.bigBlocks = new DocumentBlock[0];
                this.path = path;
                this.name = name;
                this.size = size;
            }

            private void InitBlock(POIFSDocument enclosingInstance)
            {
                this.enclosingInstance = enclosingInstance;
            }

            public event POIFSWriterEventHandler BeforeWriting;

            protected virtual void OnBeforeWriting(POIFSWriterEventArgs e)
            {
                if (BeforeWriting != null)
                {
                    BeforeWriting(this, e);
                }
            }

            internal virtual void WriteBlocks(Stream stream)
            {
                if (this.Valid)
                {
                    if (this.BeforeWriting != null)
                    {
                        POIFSDocumentWriter stream2 = new POIFSDocumentWriter(stream, this.size);
                        OnBeforeWriting(new POIFSWriterEventArgs(stream2, this.path, this.name, this.size));
                        stream2.WriteFiller(this.CountBlocks * POIFSConstants.BIG_BLOCK_SIZE, DocumentBlock.FillByte);
                    }
                    else
                    {
                        for (int i = 0; i < this.bigBlocks.Length; i++)
                        {
                            this.bigBlocks[i].WriteBlocks(stream);
                        }
                    }
                }
            }

            internal virtual DocumentBlock[] Blocks
            {
                get
                {
                    if (this.Valid && (this.BeforeWriting != null))
                    {
                        MemoryStream stream = new MemoryStream(this.size);
                        POIFSDocumentWriter stream2 = new POIFSDocumentWriter(stream, this.size);
                        //OnBeforeWriting(new POIFSWriterEventArgs(stream2, this.path, this.name, this.size));
                        this.bigBlocks = DocumentBlock.Convert(stream.ToArray(), this.size);
                    }
                    return this.bigBlocks;
                }
            }

            internal virtual int CountBlocks
            {
                get
                {
                    int num = 0;
                    if (!this.Valid)
                    {
                        return num;
                    }
                    if (this.BeforeWriting != null)
                    {
                        return (((this.size + POIFSConstants.BIG_BLOCK_SIZE) - 1) / POIFSConstants.BIG_BLOCK_SIZE);
                    }
                    return this.bigBlocks.Length;
                }
            }

            public POIFSDocument Enclosing_Instance
            {
                get
                {
                    return this.enclosingInstance;
                }
            }

            internal virtual bool Valid
            {
                get
                {
                    return ((this.bigBlocks.Length > 0) || (this.BeforeWriting != null));
                }
            }
        }

        internal class SmallBlockStore : IDisposable
        {
            private POIFSDocument enclosingInstance;
            private string name;
            private POIFSDocumentPath path;
            private int size;
            private SmallDocumentBlock[] smallBlocks;

            public void Dispose()
            {
                enclosingInstance = null;
                path = null;
                smallBlocks = null;
                name = null;
            }

            internal SmallBlockStore(POIFSDocument enclosingInstance, object[] blocks)
            {
                this.InitBlock(enclosingInstance);
                this.smallBlocks = new SmallDocumentBlock[blocks.Length];
                for (int i = 0; i < blocks.Length; i++)
                {
                    this.smallBlocks[i] = (SmallDocumentBlock)blocks[i];
                }
                this.path = null;
                this.name = null;
                this.size = -1;
            }

            internal SmallBlockStore(POIFSDocument enclosingInstance, POIFSDocumentPath path, string name, int size)
            {
                this.InitBlock(enclosingInstance);
                this.smallBlocks = new SmallDocumentBlock[0];
                this.path = path;
                this.name = name;
                this.size = size;
            }

            private void InitBlock(POIFSDocument enclosingInstance)
            {
                this.enclosingInstance = enclosingInstance;
            }

            public event POIFSWriterEventHandler BeforeWriting;

            protected virtual void OnBeforeWriting(POIFSWriterEventArgs e)
            {
                if (BeforeWriting != null)
                {
                    BeforeWriting(this, e);
                }
            }

            internal virtual BlockWritable[] Blocks
            {
                get
                {
                    if (this.Valid && (this.BeforeWriting != null))
                    {
                        MemoryStream stream = new MemoryStream(this.size);
                        POIFSDocumentWriter stream2 = new POIFSDocumentWriter(stream, this.size);
                        OnBeforeWriting(new POIFSWriterEventArgs(stream2,this.path,this.name,this.size));
                        this.smallBlocks = SmallDocumentBlock.Convert(stream.ToArray(), this.size);
                    }
                    return this.smallBlocks;
                }
            }

            public POIFSDocument Enclosing_Instance
            {
                get
                {
                    return this.enclosingInstance;
                }
            }

            internal virtual bool Valid
            {
                get
                {
                    return ((this.smallBlocks.Length > 0) || (this.BeforeWriting != null));
                }
            }
        }
    }
}
