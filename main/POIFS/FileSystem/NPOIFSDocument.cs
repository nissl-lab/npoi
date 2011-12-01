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
using NPOI.POIFS.dev.POIFSViewable;
using NPOI.POIFS.property.DocumentProperty;
using NPOI.util.HexDump;
using NPOI.util.IOUtils;

/**
 * This class manages a document in the NIO POIFS filesystem.
 * This is the {@link NPOIFSFileSystem} version.
 */
public class NPOIFSDocument : POIFSViewable {
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

      if(property.GetSize() < POIFSConstants.BIG_BLOCK_MINIMUM_DOCUMENT_SIZE) {
         _stream = new NPOIFSStream(_filesystem.GetMiniStore(), property.GetStartBlock());
         _block_size = _filesystem.GetMiniStore().GetBlockStoreBlockSize();
      } else {
         _stream = new NPOIFSStream(_filesystem, property.GetStartBlock());
         _block_size = _filesystem.GetBlockStoreBlockSize();
      }
   }

   /**
    * Constructor for a new Document
    *
    * @param name the name of the POIFSDocument
    * @param stream the InputStream we read data from
    */
   public NPOIFSDocument(String name, NPOIFSFileSystem filesystem, InputStream stream) 
       
   {
      this._filesystem = filesystem;

      // Buffer the contents into memory. This is a bit icky...
      // TODO Replace with a buffer up to the mini stream size, then streaming write
      byte[] contents;
      if(stream is MemoryStream) {
         MemoryStream bais = (MemoryStream)stream;
         contents = new byte[bais.available()];
         bais.Read(contents);
      } else {
         MemoryStream baos = new MemoryStream();
         IOUtils.copy(stream, baos);
         contents = baos.ToArray();
      }

      // Do we need to store as a mini stream or a full one?
      if(contents.Length <= POIFSConstants.BIG_BLOCK_MINIMUM_DOCUMENT_SIZE) {
         _stream = new NPOIFSStream(filesystem.GetMiniStore());
         _block_size = _filesystem.GetMiniStore().GetBlockStoreBlockSize();
      } else {
         _stream = new NPOIFSStream(filesystem);
         _block_size = _filesystem.GetBlockStoreBlockSize();
      }

      // Store it
      _stream.updateContents(contents);

      // And build the property for it
      this._property = new DocumentProperty(name, contents.Length);
      _property.SetStartBlock(_stream.GetStartBlock());     
   }
   
   int GetDocumentBlockSize() {
      return _block_size;
   }
   
   Iterator<ByteBuffer> GetBlockIterator() {
      if(getSize() > 0) {
         return _stream.GetBlockIterator();
      } else {
         List<ByteBuffer> empty = Collections.emptyList();
         return empty.iterator();
      }
   }

   /**
    * @return size of the document
    */
   public int GetSize() {
      return _property.GetSize();
   }

   /**
    * @return the instance's DocumentProperty
    */
   DocumentProperty GetDocumentProperty() {
      return _property;
   }

   /**
    * Get an array of objects, some of which may implement POIFSViewable
    *
    * @return an array of Object; may not be null, but may be empty
    */
   public Object[] GetViewableArray() {
      Object[] results = new Object[1];
      String result;

      try {
         if(getSize() > 0) {
            // Get all the data into a single array
            byte[] data = new byte[getSize()];
            int offset = 0;
            foreach(ByteBuffer buffer in _stream) {
               int length = Math.min(_block_size, data.Length-offset); 
               buffer.Get(data, offset, length);
               offset += length;
            }

            MemoryStream output = new MemoryStream();
            HexDump.dump(data, 0, output, 0);
            result = output.ToString();
         } else {
            result = "<NO DATA>";
         }
      } catch (IOException e) {
         result = e.GetMessage();
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
   public Iterator GetViewableIterator() {
      return Collections.EMPTY_LIST.iterator();
   }

   /**
    * Give viewers a hint as to whether to call GetViewableArray or
    * GetViewableIterator
    *
    * @return <code>true</code> if a viewer should call GetViewableArray,
    *		 <code>false</code> if a viewer should call GetViewableIterator
    */
   public bool preferArray() {
      return true;
   }

   /**
    * Provides a short description of the object, to be used when a
    * POIFSViewable object has not provided its contents.
    *
    * @return short description
    */
   public String GetshortDescription() {
      StringBuilder buffer = new StringBuilder();

      buffer.Append("Document: \"").Append(_property.GetName()).Append("\"");
      buffer.Append(" size = ").Append(getSize());
      return buffer.ToString();
   }
}







