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

namespace NPOI.POIFS.EventFileSystem
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;

    using NPOI.POIFS.FileSystem;
    using NPOI.POIFS.Properties;
    using NPOI.POIFS.Storage;

    /// <summary>
    /// An event-driven Reader for POIFS file systems. Users of this class
    /// first Create an instance of it, then use the RegisterListener
    /// methods to Register POIFSReaderListener instances for specific
    /// documents. Once all the listeners have been Registered, the Read()
    /// method is called, which results in the listeners being notified as
    /// their documents are Read.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class POIFSReader
    {
        public event POIFSReaderEventHandler StreamReaded;

        protected virtual void OnStreamReaded(POIFSReaderEventArgs e)
        {
            if (StreamReaded != null)
            {
                StreamReaded(this, e);
            }
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="POIFSReader"/> class.
        /// </summary>
        public POIFSReader()
        {

        }

        /// <summary>
        /// Read from an InputStream and Process the documents we Get
        /// </summary>
        /// <param name="stream">the InputStream from which to Read the data</param>
        /// <returns>POIFSDocument list</returns>
        public List<DocumentDescriptor> Read(Stream stream)
        {

            // Read the header block from the stream
            HeaderBlockReader header_block_reader = new HeaderBlockReader(stream);

            // Read the rest of the stream into blocks
            RawDataBlockList data_blocks = new RawDataBlockList(stream, header_block_reader.BigBlockSize);

            // Set up the block allocation table (necessary for the
            // data_blocks to be manageable
            new BlockAllocationTableReader(header_block_reader.BATCount,
                                           header_block_reader.BATArray,
                                           header_block_reader.XBATCount,
                                           header_block_reader.XBATIndex,
                                           data_blocks);

            // Get property table from the document
            PropertyTable properties =
                new PropertyTable(header_block_reader.PropertyStart,
                                  data_blocks);

            // Process documents
            return ProcessProperties(SmallBlockTableReader
                .GetSmallDocumentBlocks(data_blocks, properties
                    .Root, header_block_reader
                        .SBATStart), data_blocks, properties.Root
                            .Children, new POIFSDocumentPath());
        }


        /// <summary>
        /// Processes the properties.
        /// </summary>
        /// <param name="small_blocks">The small_blocks.</param>
        /// <param name="big_blocks">The big_blocks.</param>
        /// <param name="properties">The properties.</param>
        /// <param name="path">The path.</param>
        /// <returns></returns>
        private List<DocumentDescriptor> ProcessProperties(BlockList small_blocks,
                                       BlockList big_blocks,
                                       IEnumerator properties,
                                       POIFSDocumentPath path)
        {
            List<DocumentDescriptor> documents =
                new List<DocumentDescriptor>();

            while (properties.MoveNext())
            {
                Property property = (Property)properties.Current;
                String name = property.Name;

                if (property.IsDirectory)
                {
                    POIFSDocumentPath new_path = new POIFSDocumentPath(path,
                                                     new String[]
                {
                    name
                });

                    ProcessProperties(
                        small_blocks, big_blocks,
                        ((DirectoryProperty)property).Children, new_path);
                }
                else
                {
                    int startBlock = property.StartBlock;


                    int size = property.Size;
                    POIFSDocument document = null;

                    if (property.ShouldUseSmallBlocks)
                    {
                        document =
                            new POIFSDocument(name, small_blocks
                                .FetchBlocks(startBlock,-1), size);
                    }
                    else
                    {
                        document =
                            new POIFSDocument(name, big_blocks
                                .FetchBlocks(startBlock,-1), size);
                    }

                    documents.Add(
                            new DocumentDescriptor(path, name));
                    //fire event
                    OnStreamReaded(new POIFSReaderEventArgs(name, path, document));
                }
            }
            return documents;
        }
    }
}