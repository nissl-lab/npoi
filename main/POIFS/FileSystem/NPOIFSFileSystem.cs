
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



















using NPOI.POIFS.common.POIFSBigBlockSize;
using NPOI.POIFS.common.POIFSConstants;
using NPOI.POIFS.dev.POIFSViewable;
using NPOI.POIFS.nio.ByteArrayBackedDataSource;
using NPOI.POIFS.nio.DataSource;
using NPOI.POIFS.nio.FileBackedDataSource;
using NPOI.POIFS.property.DirectoryProperty;
using NPOI.POIFS.property.NPropertyTable;
using NPOI.POIFS.storage.BATBlock;
using NPOI.POIFS.storage.BlockAllocationTableReader;
using NPOI.POIFS.storage.BlockAllocationTableWriter;
using NPOI.POIFS.storage.HeaderBlock;
using NPOI.POIFS.storage.HeaderBlockConstants;
using NPOI.POIFS.storage.HeaderBlockWriter;
using NPOI.POIFS.storage.BATBlock.BATBlockAndIndex;
using NPOI.util.CloseIgnoringInputStream;
using NPOI.util.IOUtils;
using NPOI.util.LongField;
using NPOI.util.POILogFactory;
using NPOI.util.POILogger;

/**
 * This is the main class of the POIFS system; it manages the entire
 * life cycle of the filesystem.
 * This is the new NIO version
 */

public class NPOIFSFileSystem : BlockStore
    : POIFSViewable, Closeable
{
	private static POILogger _logger =
		POILogFactory.GetLogger(NPOIFSFileSystem.class);

    /**
     * Convenience method for clients that want to avoid the auto-close behaviour of the constructor.
     */
    public static InputStream CreateNonClosingInputStream(InputStream is) {
       return new CloseIgnoringInputStream(is);
    }
   
    private NPOIFSMiniStore _mini_store;
    private NPropertyTable  _property_table;
    private List<BATBlock>  _xbat_blocks;
    private List<BATBlock>  _bat_blocks;
    private HeaderBlock     _header;
    private DirectoryNode   _root;
    
    private DataSource _data;
    
    /**
     * What big block size the file uses. Most files
     *  use 512 bytes, but a few use 4096
     */
    private POIFSBigBlockSize bigBlockSize = 
       POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS;

    private NPOIFSFileSystem(bool newFS)
    {
        _header         = new HeaderBlock(bigBlockSize);
        _property_table = new NPropertyTable(_header);
        _mini_store     = new NPOIFSMiniStore(this, _property_table.GetRoot(), new List<BATBlock>(), _header);
        _xbat_blocks    = new List<BATBlock>();
        _bat_blocks     = new List<BATBlock>();
        _root           = null;
        
        if(newFS) {
           // Data needs to Initially hold just the header block,
           //  a single bat block, and an empty properties section
           _data        = new ByteArrayBackedDataSource(new byte[bigBlockSize.GetBigBlockSize()*3]);
        }
    }
    
    /**
     * Constructor, intended for writing
     */
    public NPOIFSFileSystem()
    {
       this(true);
       
        // Mark us as having a single empty BAT at offset 0
        _header.SetBATCount(1);
        _header.SetBATArray(new int[] { 0 });
        _bat_blocks.Add(BATBlock.CreateEmptyBATBlock(bigBlockSize, false));
        SetNextBlock(0, POIFSConstants.FAT_SECTOR_BLOCK);
        
        // Now associate the properties with the empty block
        _property_table.SetStartBlock(1);
        SetNextBlock(1, POIFSConstants.END_OF_CHAIN);
    }

    /**
     * Creates a POIFSFileSystem from a <tt>File</tt>. This uses less memory than
     *  creating from an <tt>InputStream</tt>. The File will be opened Read-only
     *  
     * Note that with this constructor, you will need to call {@link #close()}
     *  when you're done to have the underlying file closed, as the file is
     *  kept open during normal operation to read the data out. 
     *  
     * @param file the File from which to read the data
     *
     * @exception IOException on errors Reading, or on invalid data
     */
    public NPOIFSFileSystem(File file)
         
    {
       this(file, true);
    }
    
    /**
     * Creates a POIFSFileSystem from a <tt>File</tt>. This uses less memory than
     *  creating from an <tt>InputStream</tt>.
     *  
     * Note that with this constructor, you will need to call {@link #close()}
     *  when you're done to have the underlying file closed, as the file is
     *  kept open during normal operation to read the data out. 
     *  
     * @param file the File from which to read the data
     *
     * @exception IOException on errors Reading, or on invalid data
     */
    public NPOIFSFileSystem(File file, bool ReadOnly)
         
    {
       this(
           (new RandomAccessFile(file, ReadOnly? "r" : "rw")).GetChannel(),
           true
       );
    }
    
    /**
     * Creates a POIFSFileSystem from an open <tt>FileChannel</tt>. This uses 
     *  less memory than creating from an <tt>InputStream</tt>.
     *  
     * Note that with this constructor, you will need to call {@link #close()}
     *  when you're done to have the underlying Channel closed, as the channel is
     *  kept open during normal operation to read the data out. 
     *  
     * @param channel the FileChannel from which to read the data
     *
     * @exception IOException on errors Reading, or on invalid data
     */
    public NPOIFSFileSystem(FileChannel channel)
         
    {
       this(channel, false);
    }
    
    private NPOIFSFileSystem(FileChannel channel, bool closeChannelOnError)
         
    {
       this(false);

       try {
          // Get the header
          ByteBuffer headerBuffer = ByteBuffer.allocate(POIFSConstants.SMALLER_BIG_BLOCK_SIZE);
          IOUtils.ReadFully(channel, headerBuffer);
          
          // Have the header Processed
          _header = new HeaderBlock(headerBuffer);
          
          // Now process the various entries
          _data = new FileBackedDataSource(channel);
          ReadCoreContents();
       } catch(IOException e) {
          if(closeChannelOnError) {
             channel.Close();
          }
          throw e;
       } catch(RuntimeException e) {
          // Comes from Iterators etc.
          // TODO Decide if we can handle these better whilst
          //  still sticking to the iterator contract
          if(closeChannelOnError) {
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

    public NPOIFSFileSystem(InputStream stream)
        
    {
        this(false);
        
        ReadableByteChannel channel = null;
        bool success = false;
        
        try {
           // Turn our InputStream into something NIO based
           channel = Channels.newChannel(stream);
           
           // Get the header
           ByteBuffer headerBuffer = ByteBuffer.allocate(POIFSConstants.SMALLER_BIG_BLOCK_SIZE);
           IOUtils.ReadFully(channel, headerBuffer);
           
           // Have the header Processed
           _header = new HeaderBlock(headerBuffer);
           
           // Sanity check the block count
           BlockAllocationTableReader.sanityCheckBlockCount(_header.GetBATCount());
   
           // We need to buffer the whole file into memory when
           //  working with an InputStream.
           // The max possible size is when each BAT block entry is used
           int maxSize = BATBlock.calculateMaximumSize(_header); 
           ByteBuffer data = ByteBuffer.allocate(maxSize);
           // Copy in the header
           headerBuffer.position(0);
           data.Put(headerBuffer);
           data.position(headerBuffer.capacity());
           // Now read the rest of the stream
           IOUtils.ReadFully(channel, data);
           success = true;
           
           // Turn it into a DataSource
           _data = new ByteArrayBackedDataSource(data.array(), data.position());
        } finally {
           // As per the constructor contract, always close the stream
           if(channel != null)
              channel.Close();
           closeInputStream(stream, success);
        }
        
        // Now process the various entries
        ReadCoreContents();
    }
    /**
     * @param stream the stream to be closed
     * @param success <code>false</code> if an exception is currently being thrown in the calling method
     */
    private void closeInputStream(InputStream stream, bool success) {
        try {
            stream.Close();
        } catch (IOException e) {
            if(success) {
                throw new RuntimeException(e);
            }
            // else not success? Try block did not complete normally
            // just print stack trace and leave original ex to be thrown
            e.printStackTrace();
        }
    }

    /**
     * Checks that the supplied InputStream (which MUST
     *  support mark and Reset, or be a PushbackInputStream)
     *  has a POIFS (OLE2) header at the start of it.
     * If your InputStream does not support mark / Reset,
     *  then wrap it in a PushBackInputStream, then be
     *  sure to always use that, and not the original!
     * @param inp An InputStream which supports either mark/reset, or is a PushbackInputStream
     */
    public static bool HasPOIFSHeader(InputStream inp)  {
        // We want to peek at the first 8 bytes
        inp.mark(8);

        byte[] header = new byte[8];
        IOUtils.ReadFully(inp, header);
        LongField signature = new LongField(HeaderBlockConstants._signature_offset, header);

        // Wind back those 8 bytes
        if(inp is PushbackInputStream) {
            PushbackInputStream pin = (PushbackInputStream)inp;
            pin.unread(header);
        } else {
            inp.Reset();
        }

        // Did it match the signature?
        return (signature.Get() == HeaderBlockConstants._signature);
    }
    
    /**
     * Read and process the PropertiesTable and the
     *  FAT / XFAT blocks, so that we're Ready to
     *  work with the file
     */
    private void ReadCoreContents()  {
       // Grab the block size
       bigBlockSize = _header.GetBigBlockSize();
       
       // Each block should only ever be used by one of the
       //  FAT, XFAT or Property Table. Ensure it does
       ChainLoopDetector loopDetector = GetChainLoopDetector();
       
       // Read the FAT blocks
       foreach(int fatAt in _header.GetBATArray()) {
          ReadBAT(fatAt, loopDetector);
       }
       
       // Now read the XFAT blocks, and the FATs within them
       BATBlock xfat; 
       int nextAt = _header.GetXBATIndex();
       for(int i=0; i<_header.GetXBATCount(); i++) {
          loopDetector.claim(nextAt);
          ByteBuffer fatData = GetBlockAt(nextAt);
          xfat = BATBlock.CreateBATBlock(bigBlockSize, fatData);
          xfat.SetOurBlockIndex(nextAt);
          nextAt = xfat.GetValueAt(bigBlockSize.GetXBATEntriesPerBlock());
          _xbat_blocks.Add(xfat);
          
          for(int j=0; j<bigBlockSize.GetXBATEntriesPerBlock(); j++) {
             int fatAt = xfat.GetValueAt(j);
             if(fatAt == POIFSConstants.UNUSED_BLOCK) break;
             ReadBAT(fatAt, loopDetector);
          }
       }
       
       // We're now able to load steams
       // Use this to read in the properties
       _property_table = new NPropertyTable(_header, this);
       
       // Finally read the Small Stream FAT (SBAT) blocks
       BATBlock sfat;
       List<BATBlock> sbats = new List<BATBlock>();
       _mini_store     = new NPOIFSMiniStore(this, _property_table.GetRoot(), sbats, _header);
       nextAt = _header.GetSBATStart();
       for(int i=0; i<_header.GetSBATCount(); i++) {
          loopDetector.claim(nextAt);
          ByteBuffer fatData = GetBlockAt(nextAt);
          sfat = BATBlock.CreateBATBlock(bigBlockSize, fatData);
          sfat.SetOurBlockIndex(nextAt);
          sbats.Add(sfat);
          nextAt = GetNextBlock(nextAt);  
       }
    }
    private void ReadBAT(int batAt, ChainLoopDetector loopDetector)  {
       loopDetector.claim(batAt);
       ByteBuffer fatData = GetBlockAt(batAt);
       BATBlock bat = BATBlock.CreateBATBlock(bigBlockSize, fatData);
       bat.SetOurBlockIndex(batAt);
       _bat_blocks.Add(bat);
    }
    private BATBlock CreateBAT(int offset, bool IsBAT)  {
       // Create a new BATBlock
       BATBlock newBAT = BATBlock.CreateEmptyBATBlock(bigBlockSize, !isBAT);
       newBAT.SetOurBlockIndex(offset);
       // Ensure there's a spot in the file for it
       ByteBuffer buffer = ByteBuffer.allocate(bigBlockSize.GetBigBlockSize());
       int WriteTo = (1+offset) * bigBlockSize.GetBigBlockSize(); // Header isn't in BATs
       _data.Write(buffer, WriteTo);
       // All done
       return newBAT;
    }
    
    /**
     * Load the block at the given offset.
     */
    protected ByteBuffer GetBlockAt(final int offset)  {
       // The header block doesn't count, so add one
       long startAt = (offset+1) * bigBlockSize.GetBigBlockSize();
       return _data.Read(bigBlockSize.GetBigBlockSize(), startAt);
    }
    
    /**
     * Load the block at the given offset, 
     *  extending the file if needed
     */
    protected ByteBuffer CreateBlockIfNeeded(final int offset)  {
       try {
          return GetBlockAt(offset);
       } catch(IndexOutOfBoundsException e) {
          // The header block doesn't count, so add one
          long startAt = (offset+1) * bigBlockSize.GetBigBlockSize();
          // Allocate and write
          ByteBuffer buffer = ByteBuffer.allocate(getBigBlockSize());
          _data.Write(buffer, startAt);
          // Retrieve the properly backed block
          return GetBlockAt(offset);
       }
    }
    
    /**
     * Returns the BATBlock that handles the specified offset,
     *  and the relative index within it
     */
    protected BATBlockAndIndex GetBATBlockAndIndex(final int offset) {
       return BATBlock.GetBATBlockAndIndex(
             offset, _header, _bat_blocks
       );
    }
    
    /**
     * Works out what block follows the specified one.
     */
    protected int GetNextBlock(final int offset) {
       BATBlockAndIndex bai = GetBATBlockAndIndex(offset);
       return bai.GetBlock().GetValueAt( bai.GetIndex() );
    }
    
    /**
     * Changes the record of what block follows the specified one.
     */
    protected void SetNextBlock(final int offset, int nextBlock) {
       BATBlockAndIndex bai = GetBATBlockAndIndex(offset);
       bai.GetBlock().SetValueAt(
             bai.GetIndex(), nextBlock
       );
    }
    
    /**
     * Finds a free block, and returns its offset.
     * This method will extend the file if needed, and if doing
     *  so, allocate new FAT blocks to Address the extra space.
     */
    protected int GetFreeBlock()  {
       // First up, do we have any spare ones?
       int offset = 0;
       for(int i=0; i<_bat_blocks.Count; i++) {
          int numSectors = bigBlockSize.GetBATEntriesPerBlock();

          // Check this one
          BATBlock bat = _bat_blocks.Get(i);
          if(bat.HasFreeSectors()) {
             // Claim one of them and return it
             for(int j=0; j<numSectors; j++) {
                int batValue = bat.GetValueAt(j);
                if(batValue == POIFSConstants.UNUSED_BLOCK) {
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
       if(_header.GetBATCount() >= 109) {
          // Needs to come from an XBAT
          BATBlock xbat = null;
          foreach(BATBlock x in _xbat_blocks) {
             if(x.HasFreeSectors()) {
                xbat = x;
                break;
             }
          }
          if(xbat == null) {
             // Oh joy, we need a new XBAT too...
             xbat = CreateBAT(offset+1, false);
             xbat.SetValueAt(0, offset);
             bat.SetValueAt(1, POIFSConstants.DIFAT_SECTOR_BLOCK);
             
             // Will go one place higher as XBAT Added in
             offset++;
             
             // Chain it
             if(_xbat_blocks.Count == 0) {
                _header.SetXBATStart(offset);
             } else {
                _xbat_blocks.Get(_xbat_blocks.Count-1).SetValueAt(
                      bigBlockSize.GetXBATEntriesPerBlock(), offset
                );
             }
             _xbat_blocks.Add(xbat);
             _header.SetXBATCount(_xbat_blocks.Count);
          }
          // Allocate us in the XBAT
          for(int i=0; i<bigBlockSize.GetXBATEntriesPerBlock(); i++) {
             if(xbat.GetValueAt(i) == POIFSConstants.UNUSED_BLOCK) {
                xbat.SetValueAt(i, offset);
             }
          }
       } else {
          // Store us in the header
          int[] newBATs = new int[_header.GetBATCount()+1];
          Array.Copy(_header.GetBATArray(), 0, newBATs, 0, newBATs.Length-1);
          newBATs[newBATs.Length-1] = offset;
          _header.SetBATArray(newBATs);
       }
       _header.SetBATCount(_bat_blocks.Count);
       
       // The current offset stores us, but the next one is free
       return offset+1;
    }
    
    
    protected ChainLoopDetector GetChainLoopDetector()  {
      return new ChainLoopDetector(_data.Count);
    }

   /**
     * For unit Testing only! Returns the underlying
     *  properties table
     */
    NPropertyTable _get_property_table() {
      return _property_table;
    }
    
    /**
     * Returns the MiniStore, which performs a similar low
     *  level function to this, except for the small blocks.
     */
    public NPOIFSMiniStore GetMiniStore() {
       return _mini_store;
    }

    /**
     * add a new POIFSDocument to the FileSytem 
     *
     * @param document the POIFSDocument being Added
     */
    void AddDocument(final NPOIFSDocument document)
    {
        _property_table.AddProperty(document.GetDocumentProperty());
    }

    /**
     * add a new DirectoryProperty to the FileSystem
     *
     * @param directory the DirectoryProperty being Added
     */
    void AddDirectory(final DirectoryProperty directory)
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

    public DocumentEntry CreateDocument(final InputStream stream,
                                        String name)
        
    {
        return GetRoot().CreateDocument(name, stream);
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

    public DocumentEntry CreateDocument(final String name, int size,
                                        POIFSWriterListener Writer)
        
    {
        return GetRoot().CreateDocument(name, size, Writer);
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

    public DirectoryEntry CreateDirectory(final String name)
        
    {
        return GetRoot().CreateDirectory(name);
    }
    
    /**
     * Write the filesystem out to the open file. Will thrown an
     *  {@link ArgumentException} if opened from an 
     *  {@link InputStream}.
     * 
     * @exception IOException thrown on errors writing to the stream
     */
    public void WriteFilesystem() 
    {
       if(_data is FileBackedDataSource) {
          // Good, correct type
       } else {
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

    public void WriteFilesystem(final OutputStream stream)
        
    {
       // Have the datasource updated
       syncWithDataSource();
       
       // Now copy the contents to the stream
       _data.copyTo(stream);
    }
    
    /**
     * Has our in-memory objects write their state
     *  to their backing blocks 
     */
    private void syncWithDataSource() 
    {
       // HeaderBlock
       HeaderBlockWriter hbw = new HeaderBlockWriter(_header);
       hbw.WriteBlock( GetBlockAt(-1) );
       
       // BATs
       foreach(BATBlock bat in _bat_blocks) {
          ByteBuffer block = GetBlockAt(bat.GetOurBlockIndex());
          BlockAllocationTableWriter.WriteBlock(bat, block);
       }
       
       // SBATs
       _mini_store.syncWithDataSource();
       
       // Properties
       _property_table.Write(
             new NPOIFSStream(this, _header.GetPropertyStart())
       );
    }
    
    /**
     * Closes the FileSystem, freeing any underlying files, streams
     *  and buffers. After this, you will be unable to read or 
     *  write from the FileSystem.
     */
    public void close()  {
       _data.Close();
    }

    /**
     * read in a file and write it back out again
     *
     * @param args names of the files; arg[ 0 ] is the input file,
     *             arg[ 1 ] is the output file
     *
     * @exception IOException
     */

    public static void main(String args[])
        
    {
        if (args.Length != 2)
        {
            System.err.println(
                "two arguments required: input filename and output filename");
            System.exit(1);
        }
        FileInputStream  istream = new FileInputStream(args[ 0 ]);
        FileOutputStream ostream = new FileOutputStream(args[ 1 ]);

        new NPOIFSFileSystem(istream).WriteFilesystem(ostream);
        istream.Close();
        ostream.Close();
    }

    /**
     * Get the root entry
     *
     * @return the root entry
     */
    public DirectoryNode GetRoot()
    {
        if (_root == null) {
           _root = new DirectoryNode(_property_table.GetRoot(), this, null);
        }
        return _root;
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

    public DocumentInputStream CreateDocumentInputStream(
            String documentName)
        
    {
    	return GetRoot().CreateDocumentInputStream(documentName);
    }

    /**
     * remove an entry
     *
     * @param entry to be Removed
     */

    void Remove(EntryNode entry)
    {
        _property_table.RemoveProperty(entry.GetProperty());
    }
    
    /* ********** START begin implementation of POIFSViewable ********** */

    /**
     * Get an array of objects, some of which may implement
     * POIFSViewable
     *
     * @return an array of Object; may not be null, but may be empty
     */

    public Object [] GetViewableArray()
    {
        if (preferArray())
        {
            return (( POIFSViewable ) GetRoot()).GetViewableArray();
        }
        return new Object[ 0 ];
    }

    /**
     * Get an Iterator of objects, some of which may implement
     * POIFSViewable
     *
     * @return an Iterator; may not be null, but may have an empty
     * back end store
     */

    public Iterator GetViewableIterator()
    {
        if (!preferArray())
        {
            return (( POIFSViewable ) GetRoot()).GetViewableIterator();
        }
        return Collections.EMPTY_LIST.iterator();
    }

    /**
     * Give viewers a hint as to whether to call GetViewableArray or
     * GetViewableIterator
     *
     * @return true if a viewer should call GetViewableArray, false if
     *         a viewer should call GetViewableIterator
     */

    public bool preferArray()
    {
        return (( POIFSViewable ) GetRoot()).preferArray();
    }

    /**
     * Provides a short description of the object, to be used when a
     * POIFSViewable object has not provided its contents.
     *
     * @return short description
     */

    public String GetshortDescription()
    {
        return "POIFS FileSystem";
    }

    /* **********  END  begin implementation of POIFSViewable ********** */

    /**
     * @return The Big Block size, normally 512 bytes, sometimes 4096 bytes
     */
    public int GetBigBlockSize() {
      return bigBlockSize.GetBigBlockSize();
    }
    /**
     * @return The Big Block size, normally 512 bytes, sometimes 4096 bytes
     */
    public POIFSBigBlockSize GetBigBlockSizeDetails() {
      return bigBlockSize;
    }
    protected int GetBlockStoreBlockSize() {
       return GetBigBlockSize();
    }
}








