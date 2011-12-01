
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


namespace NPOI.POIFS.FileSystem;





using NPOI.POIFS.common.POIFSConstants;
using NPOI.POIFS.FileSystem.BlockStore.ChainLoopDetector;
using NPOI.POIFS.property.Property;
using NPOI.POIFS.storage.HeaderBlock;

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

public class NPOIFSStream : Iterable<ByteBuffer>
{
	private BlockStore blockStore;
	private int startBlock;
	
	/**
	 * Constructor for an existing stream. It's up to you
	 *  to know how to Get the start block (eg from a 
	 *  {@link HeaderBlock} or a {@link Property}) 
	 */
	public NPOIFSStream(BlockStore blockStore, int startBlock) {
	   this.blockStore = blockStore;
	   this.startBlock = startBlock;
	}
	
	/**
	 * Constructor for a new stream. A start block won't
	 *  be allocated until you begin writing to it.
	 */
	public NPOIFSStream(BlockStore blockStore) {
      this.blockStore = blockStore;
	   this.startBlock = POIFSConstants.END_OF_CHAIN;
	}
	
	/**
	 * What block does this stream start at?
	 * Will be {@link POIFSConstants#END_OF_CHAIN} for a
	 *  new stream that hasn't been written to yet.
	 */
	public int GetStartBlock() {
	   return startBlock;
	}

	/**
	 * Returns an iterator that'll supply one {@link ByteBuffer}
	 *  per block in the stream.
	 */
   public Iterator<ByteBuffer> iterator() {
      return GetBlockIterator();
   }
	
   public Iterator<ByteBuffer> GetBlockIterator() {
      if(startBlock == POIFSConstants.END_OF_CHAIN) {
         throw new InvalidOperationException(
               "Can't read from a new stream before it has been written to"
         );
      }
      return new StreamBlockByteBufferIterator(startBlock);
   }
   
   /**
    * Updates the contents of the stream to the new
    *  Set of bytes.
    * Note - if this is property based, you'll still
    *  need to update the size in the property yourself
    */
   public void updateContents(byte[] contents)  {
      // How many blocks are we going to need?
      int blockSize = blockStore.GetBlockStoreBlockSize();
      int blocks = (int)Math.ceil( ((double)contents.Length) / blockSize );
      
      // Make sure we don't encounter a loop whilst overwriting
      //  the existing blocks
      ChainLoopDetector loopDetector = blockStore.GetChainLoopDetector();
      
      // Start writing
      int prevBlock = POIFSConstants.END_OF_CHAIN;
      int nextBlock = startBlock;
      for(int i=0; i<blocks; i++) {
         int thisBlock = nextBlock;
         
         // Allocate a block if needed, otherwise figure
         //  out what the next block will be
         if(thisBlock == POIFSConstants.END_OF_CHAIN) {
            thisBlock = blockStore.GetFreeBlock();
            loopDetector.claim(thisBlock);
            
            // We're on the end of the chain
            nextBlock = POIFSConstants.END_OF_CHAIN;
            
            // Mark the previous block as carrying on to us if needed
            if(prevBlock != POIFSConstants.END_OF_CHAIN) {
               blockStore.SetNextBlock(prevBlock, thisBlock);
            }
            blockStore.SetNextBlock(thisBlock, POIFSConstants.END_OF_CHAIN);
            
            // If we've just written the first block on a 
            //  new stream, save the start block offset
            if(this.startBlock == POIFSConstants.END_OF_CHAIN) {
               this.startBlock = thisBlock;
            }
         } else {
            loopDetector.claim(thisBlock);
            nextBlock = blockStore.GetNextBlock(thisBlock);
         }
         
         // Write it
         ByteBuffer buffer = blockStore.CreateBlockIfNeeded(thisBlock);
         int startAt = i*blockSize;
         int endAt = Math.min(contents.Length - startAt, blockSize);
         buffer.Put(contents, startAt, endAt); 
         
         // Update pointers
         prevBlock = thisBlock;
      }
      int lastBlock = prevBlock;
      
      // If we're overwriting, free any remaining blocks
      NPOIFSStream toFree = new NPOIFSStream(blockStore, nextBlock);
      toFree.free(loopDetector);
      
      // Mark the end of the stream
      blockStore.SetNextBlock(lastBlock, POIFSConstants.END_OF_CHAIN);
   }
   
   // TODO Streaming write support
   // TODO  then convert fixed sized write to use streaming internally
   // TODO Append write support (probably streaming)
   
   /**
    * Frees all blocks in the stream
    */
   public void free()  {
      ChainLoopDetector loopDetector = blockStore.GetChainLoopDetector();
      free(loopDetector);
   }
   private void free(ChainLoopDetector loopDetector) {
      int nextBlock = startBlock;
      while(nextBlock != POIFSConstants.END_OF_CHAIN) {
         int thisBlock = nextBlock;
         loopDetector.claim(thisBlock);
         nextBlock = blockStore.GetNextBlock(thisBlock);
         blockStore.SetNextBlock(thisBlock, POIFSConstants.UNUSED_BLOCK);
      }
      this.startBlock = POIFSConstants.END_OF_CHAIN;
   }
   
   /**
    * Class that handles a streaming read of one stream
    */
   protected class StreamBlockByteBufferIterator : Iterator<ByteBuffer> {
      private ChainLoopDetector loopDetector;
      private int nextBlock;
      
      protected StreamBlockByteBufferIterator(int firstBlock) {
         this.nextBlock = firstBlock;
         try {
            this.loopDetector = blockStore.GetChainLoopDetector();
         } catch(IOException e) {
            throw new RuntimeException(e);
         }
      }

      public bool HasNext() {
         if(nextBlock == POIFSConstants.END_OF_CHAIN) {
            return false;
         }
         return true;
      }

      public ByteBuffer next() {
         if(nextBlock == POIFSConstants.END_OF_CHAIN) {
            throw new IndexOutOfBoundsException("Can't read past the end of the stream");
         }
         
         try {
            loopDetector.claim(nextBlock);
            ByteBuffer data = blockStore.GetBlockAt(nextBlock);
            nextBlock = blockStore.GetNextBlock(nextBlock);
            return data;
         } catch(IOException e) {
            throw new RuntimeException(e);
         }
      }

      public void Remove() {
         throw new UnsupportedOperationException();
      }
   }
}








