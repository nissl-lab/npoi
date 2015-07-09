
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
using NPOI.POIFS.FileSystem;
using System;
using System.Collections.Generic;
using System.IO;
using NPOI.Util;

namespace NPOI.POIFS.FileSystem
{

    /**
 * This handles Reading and writing a stream within a
 *  {@link NPOIFSFileSystem}. It can supply an iterator
 *  to read blocks, and way to write out to existing and
 *  new blocks.
 * Most users will want a higher level version of this, 
 *  which deals with properties to track which stream
 *  this is.
 * This only works on big block streams, it doesn't
 *  handle small block ones.
 * This uses the new NIO code
 * 
 * TODO Implement a streaming write method, and append
 */

    public class NPOIFSStream : IEnumerable<ByteBuffer>
    {
        private BlockStore blockStore;
        private int startBlock;
        private MemoryStream outStream;

        /**
         * Constructor for an existing stream. It's up to you
         *  to know how to Get the start block (eg from a 
         *  {@link HeaderBlock} or a {@link Property}) 
         */
        public NPOIFSStream(BlockStore blockStore, int startBlock)
        {
            this.blockStore = blockStore;
            this.startBlock = startBlock;
        }

        /**
         * Constructor for a new stream. A start block won't
         *  be allocated until you begin writing to it.
         */
        public NPOIFSStream(BlockStore blockStore)
        {
            this.blockStore = blockStore;
            this.startBlock = POIFSConstants.END_OF_CHAIN;
        }

        /**
         * What block does this stream start at?
         * Will be {@link POIFSConstants#END_OF_CHAIN} for a
         *  new stream that hasn't been written to yet.
         */
        public int GetStartBlock()
        {
            return startBlock;
        }


        #region IEnumerable<byte[]> Members

        //public IEnumerator<byte[]> GetEnumerator()
        public IEnumerator<ByteBuffer> GetEnumerator()
        {
            return GetBlockIterator();
        }

        #endregion


        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetBlockIterator();
        }

        #endregion

        /**
     * Returns an iterator that'll supply one {@link ByteBuffer}
     *  per block in the stream.
     */
        public IEnumerator<ByteBuffer> GetBlockIterator()
        {
            if (startBlock == POIFSConstants.END_OF_CHAIN)
            {
                throw new InvalidOperationException(
                      "Can't read from a new stream before it has been written to"
                );
            }
            return new StreamBlockByteBufferIterator(this, startBlock);
        }

        /**
         * Updates the contents of the stream to the new
         *  Set of bytes.
         * Note - if this is property based, you'll still
         *  need to update the size in the property yourself
         */
        public void UpdateContents(byte[] contents)
        {
            Stream os = GetOutputStream();
            os.Write(contents, 0, contents.Length);
            os.Close();
        }

        public Stream GetOutputStream()
        {
            if (outStream == null)
            {
                outStream = new StreamBlockByteBuffer(this);
            }
            return outStream;
        }

        // TODO Streaming write support
        // TODO  then convert fixed sized write to use streaming internally
        // TODO Append write support (probably streaming)

        /**
         * Frees all blocks in the stream
         */
        public void Free()
        {
            ChainLoopDetector loopDetector = blockStore.GetChainLoopDetector();
            Free(loopDetector);
        }
        internal void Free(ChainLoopDetector loopDetector)
        {
            int nextBlock = startBlock;
            while (nextBlock != POIFSConstants.END_OF_CHAIN)
            {
                int thisBlock = nextBlock;
                loopDetector.Claim(thisBlock);
                nextBlock = blockStore.GetNextBlock(thisBlock);
                blockStore.SetNextBlock(thisBlock, POIFSConstants.UNUSED_BLOCK);
            }
            this.startBlock = POIFSConstants.END_OF_CHAIN;
        }

        /*
         * Class that handles a streaming read of one stream
         */

        public class StreamBlockByteBuffer : MemoryStream
        {
            byte[] oneByte = new byte[1];
            ByteBuffer buffer;
            // Make sure we don't encounter a loop whilst overwriting
            // the existing blocks
            ChainLoopDetector loopDetector;
            int prevBlock, nextBlock;
            NPOIFSStream pStream;

            protected internal StreamBlockByteBuffer(NPOIFSStream pStream)
            {
                this.pStream = pStream;
                loopDetector = pStream.blockStore.GetChainLoopDetector();
                prevBlock = POIFSConstants.END_OF_CHAIN;
                nextBlock = pStream.startBlock;
            }

            protected void CreateBlockIfNeeded()
            {
                if (buffer != null && buffer.HasRemaining()) return;

                int thisBlock = nextBlock;

                // Allocate a block if needed, otherwise figure
                //  out what the next block will be
                if (thisBlock == POIFSConstants.END_OF_CHAIN)
                {
                    thisBlock = pStream.blockStore.GetFreeBlock();
                    loopDetector.Claim(thisBlock);

                    // We're on the end of the chain
                    nextBlock = POIFSConstants.END_OF_CHAIN;

                    // Mark the previous block as carrying on to us if needed
                    if (prevBlock != POIFSConstants.END_OF_CHAIN)
                    {
                        pStream.blockStore.SetNextBlock(prevBlock, thisBlock);
                    }
                    pStream.blockStore.SetNextBlock(thisBlock, POIFSConstants.END_OF_CHAIN);

                    // If we've just written the first block on a 
                    //  new stream, save the start block offset
                    if (pStream.startBlock == POIFSConstants.END_OF_CHAIN)
                    {
                        pStream.startBlock = thisBlock;
                    }
                }
                else
                {
                    loopDetector.Claim(thisBlock);
                    nextBlock = pStream.blockStore.GetNextBlock(thisBlock);
                }

                buffer = pStream.blockStore.CreateBlockIfNeeded(thisBlock);

                // Update pointers
                prevBlock = thisBlock;
            }

            public void Write(int b)
            {
                oneByte[0] = (byte)(b & 0xFF);
                base.Write(oneByte, 0, oneByte.Length);
            }

            public override void Write(byte[] b, int off, int len)
            {
                if ((off < 0) || (off > b.Length) || (len < 0) ||
                        ((off + len) > b.Length) || ((off + len) < 0))
                {
                    throw new IndexOutOfRangeException();
                }
                else if (len == 0)
                {
                    return;
                }

                do
                {
                    CreateBlockIfNeeded();
                    int writeBytes = Math.Min(buffer.Remaining(), len);
                    buffer.Write(b, off, writeBytes);
                    off += writeBytes;
                    len -= writeBytes;
                } while (len > 0);
            }

            public override void Close()
            {
                // If we're overwriting, free any remaining blocks
                NPOIFSStream toFree = new NPOIFSStream(pStream.blockStore, nextBlock);
                toFree.Free(loopDetector);

                // Mark the end of the stream
                pStream.blockStore.SetNextBlock(prevBlock, POIFSConstants.END_OF_CHAIN);

                base.Close();
            }
        }

        public class StreamBlockByteBufferIterator : IEnumerator<ByteBuffer>
        {
            private ChainLoopDetector loopDetector;
            private int nextBlock;
            //private BlockStore blockStore;
            private ByteBuffer current;
            private NPOIFSStream pStream;

            public StreamBlockByteBufferIterator(NPOIFSStream pStream, int firstBlock)
            {
                this.pStream = pStream;
                this.nextBlock = firstBlock;
                try
                {
                    this.loopDetector = pStream.blockStore.GetChainLoopDetector();
                }
                catch (IOException e)
                {
                    //throw new System.RuntimeException(e);
                    throw new Exception(e.Message);
                }
            }

            public bool HasNext()
            {
                if (nextBlock == POIFSConstants.END_OF_CHAIN)
                {
                    return false;
                }
                return true;
            }

            public ByteBuffer Next()
            {
                if (nextBlock == POIFSConstants.END_OF_CHAIN)
                {
                    throw new IndexOutOfRangeException("Can't read past the end of the stream");
                }

                try
                {
                    loopDetector.Claim(nextBlock);
                    //byte[] data = blockStore.GetBlockAt(nextBlock);
                    ByteBuffer data = pStream.blockStore.GetBlockAt(nextBlock);
                    nextBlock = pStream.blockStore.GetNextBlock(nextBlock);
                    return data;
                }
                catch (IOException e)
                {
                    throw new RuntimeException(e.Message);
                }
            }

            public void Remove()
            {
                //throw new UnsupportedOperationException();
                throw new NotImplementedException("Unsupported Operations!");
            }

            public ByteBuffer Current
            {
                get { return current; }
            }

            Object System.Collections.IEnumerator.Current
            {
                get { return current; }
            }

            void System.Collections.IEnumerator.Reset()
            {
                throw new NotImplementedException();
            }

            bool System.Collections.IEnumerator.MoveNext()
            {
                if (nextBlock == POIFSConstants.END_OF_CHAIN)
                {
                    return false;
                }

                try
                {

                    loopDetector.Claim(nextBlock);
                    // byte[] data = blockStore.GetBlockAt(nextBlock);
                    current = pStream.blockStore.GetBlockAt(nextBlock);
                    nextBlock = pStream.blockStore.GetNextBlock(nextBlock);

                    return true;
                }
                catch (IOException)
                {
                    return false;
                }
            }

            public void Dispose()
            {
            }
        }
    }

    //public class StreamBlockByteBufferIterator : IEnumerator<byte[]>
    

    
}