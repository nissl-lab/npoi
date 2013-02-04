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
using System.Collections;
using System.IO;

using NPOI.POIFS.Properties;
using NPOI.POIFS.Dev;
using NPOI.POIFS.Storage;
using NPOI.POIFS.EventFileSystem;
using NPOI.POIFS.Common;
using NPOI.Util;


namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// This is the main class of the POIFS system; it manages the entire
    /// life cycle of the filesystem.
    /// @author Marc Johnson (mjohnson at apache dot org) 
    /// </summary>
    [Serializable]
    public class POIFSFileSystem : POIFSViewable
    {
        //private static POILogger _logger =
        //        POILogFactory.GetLogger(typeof(POIFSFileSystem));
        /// <summary>
        /// Convenience method for clients that want to avoid the auto-Close behaviour of the constructor.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <example>
        /// A convenience method (
        /// CreateNonClosingInputStream()) has been provided for this purpose:
        /// StreamwrappedStream = POIFSFileSystem.CreateNonClosingInputStream(is);
        /// HSSFWorkbook wb = new HSSFWorkbook(wrappedStream);
        /// is.reset();
        /// doSomethingElse(is);
        /// </example>
        /// <returns></returns>
        public static Stream CreateNonClosingInputStream(Stream stream) {
            return new CloseIgnoringInputStream(stream);
        }

        private PropertyTable _property_table;
        private IList          _documents;
        private DirectoryNode _root;
        /**
 * What big block size the file uses. Most files
 *  use 512 bytes, but a few use 4096
 */
        private POIFSBigBlockSize bigBlockSize =
           POIFSConstants.SMALLER_BIG_BLOCK_SIZE_DETAILS;

        // there is nothing to dispose
        //public void Dispose()
        //{
        //    Dispose(true);
        //    GC.SuppressFinalize(this);
        //}

        //protected virtual void Dispose(bool disposing)
        //{
        //    if (disposing)
        //    {
        //    }
        //}

        /// <summary>
        /// Initializes a new instance of the <see cref="POIFSFileSystem"/> class.  intended for writing
        /// </summary>
        public POIFSFileSystem()
        {
            HeaderBlock headerBlock = new HeaderBlock(bigBlockSize);
            _property_table = new PropertyTable(headerBlock);
            _documents      = new ArrayList();
            _root           = null;
        }

        /// <summary>
        /// Create a POIFSFileSystem from an Stream. Normally the stream is Read until
        /// EOF.  The stream is always Closed.  In the unlikely case that the caller has such a stream and
        /// needs to use it after this constructor completes, a work around is to wrap the
        /// stream in order to trap the Close() call.  
        /// </summary>
        /// <param name="stream">the Streamfrom which to Read the data</param>
        public POIFSFileSystem(Stream stream)
            : this()
        {
            bool success = false;

            HeaderBlock header_block_reader;
            RawDataBlockList data_blocks;
            try
            {
                // Read the header block from the stream
                header_block_reader = new HeaderBlock(stream);
                bigBlockSize = header_block_reader.BigBlockSize;

                // Read the rest of the stream into blocks
                data_blocks = new RawDataBlockList(stream, bigBlockSize);
                success = true;
            }
            finally
            {
                CloseInputStream(stream, success);
            }


            // Set up the block allocation table (necessary for the
            // data_blocks to be manageable
            new BlockAllocationTableReader(header_block_reader.BigBlockSize,
                                           header_block_reader.BATCount,
                                           header_block_reader.BATArray,
                                           header_block_reader.XBATCount,
                                           header_block_reader.XBATIndex,
                                           data_blocks);

            // Get property table from the document
            PropertyTable properties = new PropertyTable(header_block_reader, data_blocks);

            // init documents
            ProcessProperties(SmallBlockTableReader.GetSmallDocumentBlocks(bigBlockSize, data_blocks, properties.Root, header_block_reader.SBATStart),
                                data_blocks, properties.Root.Children, null, header_block_reader.PropertyStart);

            // For whatever reason CLSID of root is always 0.
            Root.StorageClsid = (properties.Root.StorageClsid);
        }
        /**
         * @param stream the stream to be Closed
         * @param success <c>false</c> if an exception is currently being thrown in the calling method
         */
        private void CloseInputStream(Stream stream, bool success) {
            
            if(stream is MemoryStream) {
                String msg = "POIFS is closing the supplied input stream of type (" 
                        + stream.GetType().Name + ") which supports mark/reset.  "
                        + "This will be a problem for the caller if the stream will still be used.  "
                        + "If that is the case the caller should wrap the input stream to avoid this Close logic.  "
                        + "This warning is only temporary and will not be present in future versions of POI.";
                //_logger.Log(POILogger.WARN, msg);
            }
            try {
                stream.Close();
            } catch (IOException) {
                if(success) {
                    throw;
                }
                // else not success? Try block did not complete normally 
                // just print stack trace and leave original ex to be thrown
                //e.StackTrace;
            }
        }

        /// <summary>
        /// Checks that the supplied Stream(which MUST
        /// support mark and reset, or be a PushbackInputStream)
        /// has a POIFS (OLE2) header at the start of it.
        /// If your Streamdoes not support mark / reset,
        /// then wrap it in a PushBackInputStream, then be
        /// sure to always use that, and not the original!
        /// </summary>
        /// <param name="inp">An Streamwhich supports either mark/reset, or is a PushbackStream</param>
        /// <returns>
        /// 	<c>true</c> if [has POIFS header] [the specified inp]; otherwise, <c>false</c>.
        /// </returns>
        public static bool HasPOIFSHeader(Stream inp){

            byte[] header = new byte[8];
            IOUtils.ReadFully(inp, header);
            LongField signature = new LongField(HeaderBlockConstants._signature_offset, header);            
            
            return (signature.Value == HeaderBlockConstants._signature);
        }

        /// <summary>
        /// Create a new document to be Added to the root directory
        /// </summary>
        /// <param name="stream"> the Streamfrom which the document's data will be obtained</param>
        /// <param name="name">the name of the new POIFSDocument</param>
        /// <returns>the new DocumentEntry</returns>
        public DocumentEntry CreateDocument(Stream stream,
                                            String name)
        {
            return this.Root.CreateDocument(name, stream);
        }

        /// <summary>
        /// Create a new DocumentEntry in the root entry; the data will be
        /// provided later
        /// </summary>
        /// <param name="name">the name of the new DocumentEntry</param>
        /// <param name="size">the size of the new DocumentEntry</param>
        /// <param name="writer">the Writer of the new DocumentEntry</param>
        /// <returns>the new DocumentEntry</returns>
        public DocumentEntry CreateDocument(String name, int size,
            /*POIFSWriterEventHandler*/ POIFSWriterListener writer) //Leon
        {
            return this.Root.CreateDocument(name, size, writer);
        }

        /// <summary>
        /// Create a new DirectoryEntry in the root directory
        /// </summary>
        /// <param name="name">the name of the new DirectoryEntry</param>
        /// <returns>the new DirectoryEntry</returns>
        public DirectoryEntry CreateDirectory(String name)
        {
            return this.Root.CreateDirectory(name);
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
            return Root.CreateDocumentInputStream(documentName);
        }

        /// <summary>
        /// Writes the file system.
        /// </summary>
        /// <param name="stream">the OutputStream to which the filesystem will be
        /// written</param>
        public void WriteFileSystem(Stream stream)
        {

            // Get the property table Ready
            _property_table.PreWrite();

            // Create the small block store, and the SBAT
            SmallBlockTableWriter      sbtw       =
                new SmallBlockTableWriter(bigBlockSize,_documents, _property_table.Root);

            // Create the block allocation table
            BlockAllocationTableWriter bat        =
                new BlockAllocationTableWriter(bigBlockSize);

            // Create a list of BATManaged objects: the documents plus the
            // property table and the small block table
            ArrayList bm_objects = new ArrayList();

            bm_objects.AddRange(_documents);
            bm_objects.Add(_property_table);
            bm_objects.Add(sbtw);
            bm_objects.Add(sbtw.SBAT);

            // walk the list, allocating space for each and assigning each
            // a starting block number
            IEnumerator iter = bm_objects.GetEnumerator();

            while (iter.MoveNext())
            {
                BATManaged bmo         = ( BATManaged ) iter.Current;
                int        block_count = bmo.CountBlocks;

                if (block_count != 0)
                {
                    bmo.StartBlock=bat.AllocateSpace(block_count);
                }
                else
                {

                    // Either the BATManaged object is empty or its data
                    // is composed of SmallBlocks; in either case,
                    // allocating space in the BAT is inappropriate
                }
            }

            // allocate space for the block allocation table and take its
            // starting block
            int               batStartBlock       = bat.CreateBlocks();

            // Get the extended block allocation table blocks
            HeaderBlockWriter header_block_Writer = new HeaderBlockWriter(bigBlockSize);
            BATBlock[] xbat_blocks =
                header_block_Writer.SetBATBlocks(bat.CountBlocks,
                                                    batStartBlock);

            // Set the property table start block
            header_block_Writer.PropertyStart=_property_table.StartBlock;

            // Set the small block allocation table start block
            header_block_Writer.SBATStart=sbtw.SBAT.StartBlock;

            // Set the small block allocation table block count
            header_block_Writer.SBATBlockCount=sbtw.SBATBlockCount;

            // the header is now properly initialized. Make a list of
            // Writers (the header block, followed by the documents, the
            // property table, the small block store, the small block
            // allocation table, the block allocation table, and the
            // extended block allocation table blocks)
            ArrayList Writers = new ArrayList();

            Writers.Add(header_block_Writer);
            Writers.AddRange(_documents);
            Writers.Add(_property_table);
            Writers.Add(sbtw);
            Writers.Add(sbtw.SBAT);
            Writers.Add(bat);
            for (int j = 0; j < xbat_blocks.Length; j++)
            {
                Writers.Add(xbat_blocks[j]);
            }

            // now, Write everything out
            iter = Writers.GetEnumerator();
            while (iter.MoveNext())
            {
                BlockWritable Writer = ( BlockWritable ) iter.Current;

                Writer.WriteBlocks(stream);
            }

            Writers=null;
            iter = null;
        }

        /// <summary>
        /// Get the root entry
        /// </summary>
        /// <value>The root.</value>
        public DirectoryNode Root
        {
            get{
                if (_root == null)
                {
                    _root = new DirectoryNode(_property_table.Root, this, null);
                }
                return _root;
            }
        }

        // <summary>
        // open a document in the root entry's list of entries
        // </summary>
        // <param name="documentName">the name of the document to be opened</param>
        // <returns>a newly opened POIFSDocumentReader</returns>
        //public DocumentReader CreatePOIFSDocumentReader(
        //        String documentName)
        //{
        //    return this.Root.CreatePOIFSDocumentReader(documentName);
        //}

        /// <summary>
        /// Add a new POIFSDocument
        /// </summary>
        /// <param name="document">the POIFSDocument being Added</param>
        public void AddDocument(POIFSDocument document)
        {
            _documents.Add(document);
            _property_table.AddProperty(document.DocumentProperty);
        }

        /// <summary>
        /// Add a new DirectoryProperty
        /// </summary>
        /// <param name="directory">The directory.</param>
        public void AddDirectory(DirectoryProperty directory)
        {
            _property_table.AddProperty(directory);
        }

        /// <summary>
        /// Removes the specified entry.
        /// </summary>
        /// <param name="entry">The entry.</param>
        public void Remove(EntryNode entry)
        {
            _property_table.RemoveProperty(entry.Property);
            if (entry.IsDocumentEntry)
            {
                _documents.Remove((( DocumentNode ) entry).Document);
            }
        }

        private void ProcessProperties(BlockList small_blocks,
                                       BlockList big_blocks,
                                       IEnumerator properties,
                                       DirectoryNode dir,
                                       int headerPropertiesStartAt)
        {
            while (properties.MoveNext())
            {
                Property      property = ( Property ) properties.Current;
                String        name     = property.Name;
                DirectoryNode parent   = (dir == null)
                                         ? (( DirectoryNode ) this.Root)
                                         : dir;

                if (property.IsDirectory)
                {
                    DirectoryNode new_dir =
                        ( DirectoryNode ) parent.CreateDirectory(name);

                    new_dir.StorageClsid=property.StorageClsid ;

                    ProcessProperties(
                        small_blocks, big_blocks,
                        ((DirectoryProperty)property).Children, new_dir, headerPropertiesStartAt);
                }
                else
                {
                    int           startBlock = property.StartBlock;
                    int           size       = property.Size;
                    POIFSDocument document   = null;

                    if (property.ShouldUseSmallBlocks)
                    {
                        document =
                            new POIFSDocument(name, small_blocks
                                .FetchBlocks(startBlock, headerPropertiesStartAt), size);
                    }
                    else
                    {
                        document =
                            new POIFSDocument(name,
                                              big_blocks.FetchBlocks(startBlock,headerPropertiesStartAt),
                                              size);
                    }
                    parent.CreateDocument(document);
                }
            }
        }

        /// <summary>
        /// Get an array of objects, some of which may implement
        /// POIFSViewable        
        /// </summary>
        /// <value>an array of Object; may not be null, but may be empty</value>
        public Array ViewableArray
        {
            get{
            if (PreferArray)
            {
                    return ((POIFSViewable)this.Root).ViewableArray;
            }
            else
            {
                    return new Object[0];
            }
            }
        }

        /// <summary>
        /// Get an Iterator of objects, some of which may implement
        /// POIFSViewable
        /// </summary>
        /// <value>an Iterator; may not be null, but may have an empty
        /// back end store</value>
        public IEnumerator ViewableIterator
        {
            get{
                if (!this.PreferArray)
                {
                    return (( POIFSViewable ) this.Root).ViewableIterator;
                }
                else
                {
                    return ArrayList.ReadOnly(new ArrayList()).GetEnumerator();
                }
            }
        }

        /// <summary>
        /// Give viewers a hint as to whether to call GetViewableArray or
        /// GetViewableIterator
        /// </summary>
        /// <value><c>true</c> if a viewer should call GetViewableArray, <c>false</c> if
        /// a viewer should call GetViewableIterator </value>
        public bool PreferArray
        {
            get{return (( POIFSViewable ) this.Root).PreferArray;}
        }

        /// <summary>
        /// Provides a short description of the object, to be used when a
        /// POIFSViewable object has not provided its contents.
        /// </summary>
        /// <value>The short description.</value>
        public String ShortDescription
        {
            get{return "POIFS FileSystem";}
        }

        /// <summary>
        /// Gets The Big Block size, normally 512 bytes, sometimes 4096 bytes
        /// </summary>
        /// <value>The size of the big block.</value>
        public int BigBlockSize 
        {
            get { return bigBlockSize.GetBigBlockSize(); }
        }
    }
}
