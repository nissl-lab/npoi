
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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using NPOI.POIFS.Common;
using NPOI.POIFS.Dev;
using NPOI.POIFS.NIO;
using NPOI.POIFS.Properties;
using NPOI.POIFS.Storage;
using NPOI.Util;
using NPOI.POIFS.EventFileSystem;

namespace NPOI.POIFS.FileSystem
{


    /**
     * This is the main class of the POIFS system; it manages the entire
     * life cycle of the filesystem.
     * This is the new NIO version
     */

    public class NPOIFSFileSystem : BlockStore, POIFSViewable //, Closeable Leon
    {
        private static POILogger _logger =
                POILogFactory.GetLogger(typeof(NPOIFSFileSystem));

        /**
         * Convenience method for clients that want to avoid the auto-close behaviour of the constructor.
         */
        public static Stream CreateNonClosingInputStream(Stream stream)
        {
            return new CloseIgnoringInputStream(stream);
        }

        private NPOIFSMiniStore _mini_store;
        private NPropertyTable _property_table;
        private List<BATBlock> _xbat_blocks;
        private List<BATBlock> _bat_blocks;
        private HeaderBlock _header;
        private DirectoryNode _root;

        private DataSource _data;

        /**
         * What big block size the file uses. Most files
         *  use 512 bytes, but a few use 4096
         */
        private POIFSBigBlockSize bigBlockSize =
           POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS;

        public DataSource Data
        {
            get { return _data; }
            set { _data = value; }
        }

        private NPOIFSFileSystem(bool newFS)
        {
            _header = new HeaderBlock(bigBlockSize);
            _property_table = new NPropertyTable(_header);
            _mini_store = new NPOIFSMiniStore(this, _property_table.Root, new List<BATBlock>(), _header);
            _xbat_blocks = new List<BATBlock>();
            _bat_blocks = new List<BATBlock>();
            _root = null;

            if (newFS)
            {
                // Data needs to Initially hold just the header block,
                //  a single bat block, and an empty properties section
                _data = new ByteArrayBackedDataSource(new byte[bigBlockSize.GetBigBlockSize() * 3]);
            }
        }

        /**
         * Constructor, intended for writing
         */
        public NPOIFSFileSystem()
            : this(true)
        {

            // Mark us as having a single empty BAT at offset 0
            _header.BATCount = 1;
            _header.BATArray = new int[] { 0 };

            BATBlock bb = BATBlock.CreateEmptyBATBlock(bigBlockSize, false);
            bb.OurBlockIndex = 0;
            _bat_blocks.Add(bb);

            SetNextBlock(0, POIFSConstants.FAT_SECTOR_BLOCK);
            SetNextBlock(1, POIFSConstants.END_OF_CHAIN);

            _property_table.StartBlock = (POIFSConstants.END_OF_CHAIN);
        }

        /**
         * <p>Creates a POIFSFileSystem from a <tt>File</tt>. This uses less memory than
         *  creating from an <tt>InputStream</tt>.</p>
         *  
         * <p>Note that with this constructor, you will need to call {@link #close()}
         *  when you're done to have the underlying file closed, as the file is
         *  kept open during normal operation to read the data out.</p> 
         *  
         * @param file the File from which to read or read/write the data
         * @param readOnly whether the POIFileSystem will only be used in read-only mode
         *
         * @exception IOException on errors reading, or on invalid data
         */
        public NPOIFSFileSystem(FileInfo file, bool readOnly)
            : this(null, file, readOnly, true)
        {
            ;
        }
        /**
         * <p>Creates a POIFSFileSystem from an open <tt>FileChannel</tt>. This uses 
         *  less memory than creating from an <tt>InputStream</tt>. The stream will
        *  be used in read-only mode.</p>
         *  
         * <p>Note that with this constructor, you will need to call {@link #close()}
         *  when you're done to have the underlying Channel closed, as the channel is
         *  kept open during normal operation to read the data out.</p> 
         *  
         * @param channel the FileChannel from which to read the data
         *
         * @exception IOException on errors reading, or on invalid data
         */
        public NPOIFSFileSystem(FileStream channel)
            : this(channel, true)
        {

        }

        /**
         * <p>Creates a POIFSFileSystem from an open <tt>FileChannel</tt>. This uses 
         *  less memory than creating from an <tt>InputStream</tt>.</p>
         *  
         * <p>Note that with this constructor, you will need to call {@link #close()}
         *  when you're done to have the underlying Channel closed, as the channel is
         *  kept open during normal operation to read the data out.</p> 
         *  
         * @param channel the FileChannel from which to read or read/write the data
         * @param readOnly whether the POIFileSystem will only be used in read-only mode
         *
         * @exception IOException on errors reading, or on invalid data
         */
        public NPOIFSFileSystem(FileStream channel, bool readOnly)
            : this(channel, null, readOnly, false)
        {
            ;
        }
        public NPOIFSFileSystem(FileStream channel, FileInfo srcFile, bool readOnly, bool closeChannelOnError)
            : this(false)
        {

            try
            {
                // Initialize the datasource
                if (srcFile != null)
                {
                    //FileBackedDataSource d = new FileBackedDataSource(srcFile, readOnly);
                    channel = new FileStream(srcFile.FullName, FileMode.Open, FileAccess.Read);
                    _data = new FileBackedDataSource(channel, readOnly);
                }
                else
                {
                    _data = new FileBackedDataSource(channel, readOnly);
                }
                // Get the header
                byte[] headerBuffer = new byte[POIFSConstants.SMALLER_BIG_BLOCK_SIZE];
                IOUtils.ReadFully(channel, headerBuffer);

                // Have the header Processed
                _header = new HeaderBlock(headerBuffer);

                // Now process the various entries
                //_data = new FileBackedDataSource(channel, readOnly);
                ReadCoreContents();
                channel.Close();

            }
            catch (IOException e)
            {
                if (closeChannelOnError)
                {
                    if (channel != null)
                        channel.Close();
                }
                throw e;
            }
            catch (Exception e)
            {
                // Comes from Iterators etc.
                // TODO Decide if we can handle these better whilst
                //  still sticking to the iterator contract
                if (closeChannelOnError)
                {
                    if (channel != null)
                    channel.Close();
                }
                throw e;
            }
        }

        /**
         * Create a POIFSFileSystem from an <tt>InputStream</tt>.  Normally the stream is read until
         * EOF.  The stream is always closed.<p/>
         *
         * Some streams are usable After reaching EOF (typically those that return <code>true</code>
         * for <tt>markSupported()</tt>).  In the unlikely case that the caller has such a stream
         * <i>and</i> needs to use it After this constructor completes, a work around is to wrap the
         * stream in order to trap the <tt>close()</tt> call.  A convenience method (
         * <tt>CreateNonClosingInputStream()</tt>) has been provided for this purpose:
         * <pre>
         * InputStream wrappedStream = POIFSFileSystem.CreateNonClosingInputStream(is);
         * HSSFWorkbook wb = new HSSFWorkbook(wrappedStream);
         * is.Reset();
         * doSomethingElse(is);
         * </pre>
         * Note also the special case of <tt>MemoryStream</tt> for which the <tt>close()</tt>
         * method does nothing.
         * <pre>
         * MemoryStream bais = ...
         * HSSFWorkbook wb = new HSSFWorkbook(bais); // calls bais.Close() !
         * bais.Reset(); // no problem
         * doSomethingElse(bais);
         * </pre>
         *
         * @param stream the InputStream from which to read the data
         *
         * @exception IOException on errors Reading, or on invalid data
         */

        public NPOIFSFileSystem(Stream stream)
            : this(false)
        {

            Stream channel = null;
            bool success = false;

            try
            {
                // Turn our InputStream into something NIO based
                channel = stream;

                // Get the header
                ByteBuffer headerBuffer = ByteBuffer.CreateBuffer(POIFSConstants.SMALLER_BIG_BLOCK_SIZE);
                IOUtils.ReadFully(channel, headerBuffer.Buffer);

                // Have the header Processed
                _header = new HeaderBlock(headerBuffer);

                // Sanity check the block count
                BlockAllocationTableReader.SanityCheckBlockCount(_header.BATCount);

                // We need to buffer the whole file into memory when
                //  working with an InputStream.
                // The max possible size is when each BAT block entry is used
                long maxSize = BATBlock.CalculateMaximumSize(_header);
                if (maxSize > int.MaxValue)
                {
                    throw new ArgumentException("Unable read a >2gb file via an InputStream");
                }

                ByteBuffer data = ByteBuffer.CreateBuffer((int)maxSize);
                headerBuffer.Position = 0;
                data.Write(headerBuffer.Buffer);
                data.Position = headerBuffer.Length;

                //IOUtils.ReadFully(channel, data);
                data.Position += IOUtils.ReadFully(channel, data.Buffer, data.Position, (int)channel.Length);
                success = true;

                // Turn it into a DataSource
                _data = new ByteArrayBackedDataSource(data.Buffer, data.Position);
            }
            finally
            {
                // As per the constructor contract, always close the stream
                if (channel != null)
                    channel.Close();
                CloseInputStream(stream, success);
            }

            // Now process the various entries
            ReadCoreContents();
        }
        /**
         * @param stream the stream to be closed
         * @param success <code>false</code> if an exception is currently being thrown in the calling method
         */
        private void CloseInputStream(Stream stream, bool success)
        {
            try
            {
                stream.Close();
            }
            catch (IOException e)
            {
                if (success)
                {
                    throw new Exception(e.Message);
                }

            }
        }

        /**
         * Checks that the supplied InputStream (which MUST
         *  support mark and reset, or be a PushbackInputStream)
         *  has a POIFS (OLE2) header at the start of it.
         * If your InputStream does not support mark / reset,
         *  then wrap it in a PushBackInputStream, then be
         *  sure to always use that, and not the original!
         * @param inp An InputStream which supports either mark/reset, or is a PushbackInputStream
         */
        public static bool HasPOIFSHeader(Stream inp)
        {
            // We want to peek at the first 8 bytes
            //inp.Mark(8);

            byte[] header = new byte[8];
            IOUtils.ReadFully(inp, header);
            LongField signature = new LongField(HeaderBlockConstants._signature_offset, header);

            // Wind back those 8 bytes
            inp.Position = 0;

            // Did it match the signature?
            return (signature.Value == HeaderBlockConstants._signature);
        }

        /**
         * Checks if the supplied first 8 bytes of a stream / file
         *  has a POIFS (OLE2) header.
         */
        public static bool HasPOIFSHeader(byte[] header8Bytes)
        {
            LongField signature = new LongField(HeaderBlockConstants._signature_offset, header8Bytes);
            return (signature.Value == HeaderBlockConstants._signature);
        }


        /**
         * Read and process the PropertiesTable and the
         *  FAT / XFAT blocks, so that we're Ready to
         *  work with the file
         */
        private void ReadCoreContents()
        {
            // Grab the block size
            bigBlockSize = _header.BigBlockSize;

            // Each block should only ever be used by one of the
            //  FAT, XFAT or Property Table. Ensure it does
            ChainLoopDetector loopDetector = GetChainLoopDetector();

            // Read the FAT blocks
            foreach (int fatAt in _header.BATArray)
            {
                ReadBAT(fatAt, loopDetector);
            }
            // Work out how many FAT blocks remain in the XFATs
            int remainingFATs = _header.BATCount - _header.BATArray.Length;
       
            // Now read the XFAT blocks, and the FATs within them
            BATBlock xfat;
            int nextAt = _header.XBATIndex;
            for (int i = 0; i < _header.XBATCount; i++)
            {
                loopDetector.Claim(nextAt);
                ByteBuffer fatData = GetBlockAt(nextAt);
                xfat = BATBlock.CreateBATBlock(bigBlockSize, fatData);
                xfat.OurBlockIndex = nextAt;
                nextAt = xfat.GetValueAt(bigBlockSize.GetXBATEntriesPerBlock());
                _xbat_blocks.Add(xfat);
                // Process all the (used) FATs from this XFAT
                int xbatFATs = Math.Min(remainingFATs, bigBlockSize.GetXBATEntriesPerBlock());
                for(int j=0; j<xbatFATs; j++) 
                {
                    int fatAt = xfat.GetValueAt(j);
                    if (fatAt == POIFSConstants.UNUSED_BLOCK || fatAt == POIFSConstants.END_OF_CHAIN) break;
                    ReadBAT(fatAt, loopDetector);
                }
                remainingFATs -= xbatFATs;
            }

            // We're now able to load steams
            // Use this to read in the properties
            _property_table = new NPropertyTable(_header, this);

            // Finally read the Small Stream FAT (SBAT) blocks
            BATBlock sfat;
            List<BATBlock> sbats = new List<BATBlock>();
            _mini_store = new NPOIFSMiniStore(this, _property_table.Root, sbats, _header);
            nextAt = _header.SBATStart;
            for (int i = 0; i < _header.SBATCount; i++)
            {
                loopDetector.Claim(nextAt);
                ByteBuffer fatData = GetBlockAt(nextAt);
                sfat = BATBlock.CreateBATBlock(bigBlockSize, fatData);
                sfat.OurBlockIndex = nextAt;
                sbats.Add(sfat);
                nextAt = GetNextBlock(nextAt);
            }
        }

        private void ReadBAT(int batAt, ChainLoopDetector loopDetector)
        {
            loopDetector.Claim(batAt);
            ByteBuffer fatData = GetBlockAt(batAt);
            // byte[] fatData = GetBlockAt(batAt);
            BATBlock bat = BATBlock.CreateBATBlock(bigBlockSize, fatData);
            bat.OurBlockIndex = batAt;
            _bat_blocks.Add(bat);
        }
        private BATBlock CreateBAT(int offset, bool isBAT)
        {
            // Create a new BATBlock
            BATBlock newBAT = BATBlock.CreateEmptyBATBlock(bigBlockSize, !isBAT);
            newBAT.OurBlockIndex = offset;
            // Ensure there's a spot in the file for it
            ByteBuffer buffer = ByteBuffer.CreateBuffer(bigBlockSize.GetBigBlockSize());
            int WriteTo = (1 + offset) * bigBlockSize.GetBigBlockSize(); // Header isn't in BATs
            _data.Write(buffer, WriteTo);
            // All done
            return newBAT;
        }

        /**
         * Load the block at the given offset.
         */
        public override ByteBuffer GetBlockAt(int offset)
        {
            // The header block doesn't count, so add one
            long startAt = (offset + 1) * bigBlockSize.GetBigBlockSize();
            try
            {
                return _data.Read(bigBlockSize.GetBigBlockSize(), startAt);
            }
            catch (IndexOutOfRangeException e)
            {
                throw new IndexOutOfRangeException("Block " + offset + " not found - " + e);
            }
            
        }

        /**
         * Load the block at the given offset, 
         *  extending the file if needed
         */
        public override ByteBuffer CreateBlockIfNeeded(int offset)
        {
            try
            {
                return GetBlockAt(offset);
            }
            catch (IndexOutOfRangeException)
            {
                // The header block doesn't count, so add one
                long startAt = (offset + 1) * bigBlockSize.GetBigBlockSize();
                // Allocate and write
                ByteBuffer buffer = ByteBuffer.CreateBuffer(GetBigBlockSize());
                // byte[] buffer = new byte[GetBigBlockSize()];
                _data.Write(buffer, startAt);
                // Retrieve the properly backed block
                return GetBlockAt(offset);
            }
        }

        /**
         * Returns the BATBlock that handles the specified offset,
         *  and the relative index within it
         */
        public override BATBlockAndIndex GetBATBlockAndIndex(int offset)
        {
            return BATBlock.GetBATBlockAndIndex(offset, _header, _bat_blocks);
        }

        /**
         * Works out what block follows the specified one.
         */
        public override int GetNextBlock(int offset)
        {
            BATBlockAndIndex bai = GetBATBlockAndIndex(offset);
            return bai.Block.GetValueAt(bai.Index);
        }

        /**
         * Changes the record of what block follows the specified one.
         */
        public override void SetNextBlock(int offset, int nextBlock)
        {
            BATBlockAndIndex bai = GetBATBlockAndIndex(offset);
            bai.Block.SetValueAt(bai.Index, nextBlock);
        }

        /**
         * Finds a free block, and returns its offset.
         * This method will extend the file if needed, and if doing
         *  so, allocate new FAT blocks to Address the extra space.
         */
        public override int GetFreeBlock()
        {
            int numSectors = bigBlockSize.GetBATEntriesPerBlock();
            // First up, do we have any spare ones?
            int offset = 0;
            foreach(BATBlock temp in _bat_blocks)
            {
                if (temp.HasFreeSectors)
                {
                    // Claim one of them and return it
                    for (int j = 0; j < numSectors; j++)
                    {
                        int batValue = temp.GetValueAt(j);
                        if (batValue == POIFSConstants.UNUSED_BLOCK)
                        {
                            // Bingo
                            return offset + j;
                        }
                    }
                }

                // Move onto the next BAT
                offset += numSectors;
            }

            // If we Get here, then there aren't any free sectors
            //  in any of the BATs, so we need another BAT
            BATBlock bat = CreateBAT(offset, true);
            bat.SetValueAt(0, POIFSConstants.FAT_SECTOR_BLOCK);
            _bat_blocks.Add(bat);

            // Now store a reference to the BAT in the required place 
            if (_header.BATCount >= 109)
            {
                // Needs to come from an XBAT
                BATBlock xbat = null;
                foreach (BATBlock x in _xbat_blocks)
                {
                    if (x.HasFreeSectors)
                    {
                        xbat = x;
                        break;
                    }
                }
                if (xbat == null)
                {
                    // Oh joy, we need a new XBAT too...
                    xbat = CreateBAT(offset + 1, false);
                    // Allocate our new BAT as the first block in the XBAT
                    xbat.SetValueAt(0, offset);
                    // And allocate the XBAT in the BAT
                    bat.SetValueAt(1, POIFSConstants.DIFAT_SECTOR_BLOCK);

                    // Will go one place higher as XBAT Added in
                    offset++;

                    // Chain it
                    if (_xbat_blocks.Count == 0)
                    {
                        _header.XBATStart = offset;
                    }
                    else
                    {
                        _xbat_blocks[_xbat_blocks.Count - 1].SetValueAt(
                          bigBlockSize.GetXBATEntriesPerBlock(), offset
                    );
                    }
                    _xbat_blocks.Add(xbat);
                    _header.XBATCount = _xbat_blocks.Count;
                }
                else
                {
                    // Allocate us in the XBAT
                    for (int i = 0; i < bigBlockSize.GetXBATEntriesPerBlock(); i++)
                    {
                        if (xbat.GetValueAt(i) == POIFSConstants.UNUSED_BLOCK)
                        {
                            xbat.SetValueAt(i, offset);
                            break;
                        }
                    }
                }
            }
            else
            {
                // Store us in the header
                int[] newBATs = new int[_header.BATCount + 1];
                Array.Copy(_header.BATArray, 0, newBATs, 0, newBATs.Length - 1);
                newBATs[newBATs.Length - 1] = offset;
                _header.BATArray = newBATs;
            }
            _header.BATCount = _bat_blocks.Count;

            // The current offset stores us, but the next one is free
            return offset + 1;
        }

        protected internal long Size
        {
            get
            {
                return _data.Size;
            }
        }
        public override ChainLoopDetector GetChainLoopDetector()
        {
            return new ChainLoopDetector(_data.Size, this);
        }

        /**
          * For unit Testing only! Returns the underlying
          *  properties table
          */
        public NPropertyTable PropertyTable
        {
            get { return _property_table; }
        }

        /**
         * Returns the MiniStore, which performs a similar low
         *  level function to this, except for the small blocks.
         */
        public NPOIFSMiniStore GetMiniStore()
        {
            return _mini_store;
        }

        /**
         * add a new POIFSDocument to the FileSytem 
         *
         * @param document the POIFSDocument being Added
         */
        public void AddDocument(NPOIFSDocument document)
        {
            _property_table.AddProperty(document.DocumentProperty);
        }

        /**
         * add a new DirectoryProperty to the FileSystem
         *
         * @param directory the DirectoryProperty being Added
         */
        public void AddDirectory(DirectoryProperty directory)
        {
            _property_table.AddProperty(directory);
        }

        /**
          * Create a new document to be Added to the root directory
          *
          * @param stream the InputStream from which the document's data
          *               will be obtained
          * @param name the name of the new POIFSDocument
          *
          * @return the new DocumentEntry
          *
          * @exception IOException on error creating the new POIFSDocument
          */

        public DocumentEntry CreateDocument(Stream stream, String name)
        {
            return Root.CreateDocument(name, stream);
        }

        /**
         * create a new DocumentEntry in the root entry; the data will be
         * provided later
         *
         * @param name the name of the new DocumentEntry
         * @param size the size of the new DocumentEntry
         * @param Writer the Writer of the new DocumentEntry
         *
         * @return the new DocumentEntry
         *
         * @exception IOException
         */

        public DocumentEntry CreateDocument(String name, int size, POIFSWriterListener writer)
        {
            return Root.CreateDocument(name, size, writer);
        }

        /**
         * create a new DirectoryEntry in the root directory
         *
         * @param name the name of the new DirectoryEntry
         *
         * @return the new DirectoryEntry
         *
         * @exception IOException on name duplication
         */

        public DirectoryEntry CreateDirectory(String name)
        {
            return Root.CreateDirectory(name);
        }

        /**
         * Write the filesystem out to the open file. Will thrown an
         *  {@link ArgumentException} if opened from an 
         *  {@link InputStream}.
         * 
         * @exception IOException thrown on errors writing to the stream
         */
        public void WriteFileSystem()
        {
            if (_data is FileBackedDataSource)
            {
                // Good, correct type
            }
            else
            {
                throw new ArgumentException(
                      "POIFS opened from an inputstream, so WriteFilesystem() may " +
                      "not be called. Use WriteFilesystem(OutputStream) instead"
                );
            }
            syncWithDataSource();
        }

        /**
         * Write the filesystem out
         *
         * @param stream the OutputStream to which the filesystem will be
         *               written
         *
         * @exception IOException thrown on errors writing to the stream
         */

        public void WriteFileSystem(Stream stream)
        {

            // Have the datasource updated
            syncWithDataSource();

            // Now copy the contents to the stream
            _data.CopyTo(stream);
        }

        /**
         * Has our in-memory objects write their state
         *  to their backing blocks 
         */
        private void syncWithDataSource()
        {
            // Properties
            NPOIFSStream propStream = new NPOIFSStream(this, _header.PropertyStart);
            _property_table.PreWrite();
            _property_table.Write(propStream);
            // _header.setPropertyStart has been updated on write ...
            // HeaderBlock
            HeaderBlockWriter hbw = new HeaderBlockWriter(_header);
            hbw.WriteBlock(GetBlockAt(-1));

            // BATs
            foreach (BATBlock bat in _bat_blocks)
            {
                ByteBuffer block = GetBlockAt(bat.OurBlockIndex);
                //byte[] block = GetBlockAt(bat.OurBlockIndex);
                BlockAllocationTableWriter.WriteBlock(bat, block);
            }
            // XBats
            foreach (BATBlock bat in _xbat_blocks)
            {
                ByteBuffer block = GetBlockAt(bat.OurBlockIndex);
                BlockAllocationTableWriter.WriteBlock(bat, block);
            }
            // SBATs
            _mini_store.SyncWithDataSource();
        }

        /**
         * Closes the FileSystem, freeing any underlying files, streams
         *  and buffers. After this, you will be unable to read or 
         *  write from the FileSystem.
         */
        public void Close()
        {
            _data.Close();
        }
        /**
         * Get the root entry
         *
         * @return the root entry
         */
        public DirectoryNode Root
        {
            get
            {
                if (_root == null)
                {
                    _root = new DirectoryNode(_property_table.Root, this, null);
                }
                return _root;
            }
        }

        /**
         * open a document in the root entry's list of entries
         *
         * @param documentName the name of the document to be opened
         *
         * @return a newly opened DocumentInputStream
         *
         * @exception IOException if the document does not exist or the
         *            name is that of a DirectoryEntry
         */
        public DocumentInputStream CreateDocumentInputStream(string documentName)
        {
            return Root.CreateDocumentInputStream(documentName);
        }

        /**
         * remove an entry
         *
         * @param entry to be Removed
         */

        public void Remove(EntryNode entry)
        {
            // If it's a document, free the blocks
            if (entry is DocumentEntry) {
                NPOIFSDocument doc = new NPOIFSDocument((DocumentProperty)entry.Property, this);
                doc.Free();
            }
            _property_table.RemoveProperty(entry.Property);
        }

        /* ********** START begin implementation of POIFSViewable ********** */

        /**
         * Get an array of objects, some of which may implement
         * POIFSViewable
         *
         * @return an array of Object; may not be null, but may be empty
         */

        protected Object[] GetViewableArray()
        {
            if (PreferArray)
            {
                Array ar = ((POIFSViewable)Root).ViewableArray;
                Object[] rval = new Object[ar.Length];

                for (int i = 0; i < ar.Length; i++)
                    rval[i] = ar.GetValue(i);

                return rval;

            }
            return new Object[0];
        }

        /**
         * Get an Iterator of objects, some of which may implement
         * POIFSViewable
         *
         * @return an Iterator; may not be null, but may have an empty
         * back end store
         */

        protected IEnumerator GetViewableIterator()
        {
            if (!PreferArray)
            {
                return ((POIFSViewable)Root).ViewableIterator;
            }
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
            return "POIFS FileSystem";
        }

        /* **********  END  begin implementation of POIFSViewable ********** */

        /**
         * @return The Big Block size, normally 512 bytes, sometimes 4096 bytes
         */
        public int GetBigBlockSize()
        {
            return bigBlockSize.GetBigBlockSize();
        }
        /**
         * @return The Big Block size, normally 512 bytes, sometimes 4096 bytes
         */
        public POIFSBigBlockSize GetBigBlockSizeDetails()
        {
            return bigBlockSize;
        }
        public override int GetBlockStoreBlockSize()
        {
            return GetBigBlockSize();
        }

        #region POIFSViewable Members

        public bool PreferArray
        {
            get { return ((POIFSViewable)Root).PreferArray; }
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