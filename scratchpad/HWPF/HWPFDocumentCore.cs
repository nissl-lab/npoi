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

namespace NPOI.HWPF
{
    using NPOI.POIFS.FileSystem;
    using System;
    using System.IO;
    using NPOI.Util;
    using NPOI.HWPF.Model;
    using NPOI.HWPF.UserModel;
    using System.Text;

    /// <summary>
    /// This class holds much of the core of a Word document, but without some of the table structure information. You generally want to work with one of HWPFDocument or HWPFOldDocument
    /// </summary>
    public abstract class HWPFDocumentCore : POIDocument
    {
        protected const String STREAM_OBJECT_POOL = "ObjectPool";
        protected const String STREAM_WORD_DOCUMENT = "WordDocument";
        /** Holds OLE2 objects */
        protected ObjectPoolImpl _objectPool;

        /** The FIB */
        protected FileInformationBlock _fib;

        /** Holds styles for this document.*/
        protected StyleSheet _ss;

        /** Contains formatting properties for text*/
        protected CHPBinTable _cbt;

        /** Contains formatting properties for paragraphs*/
        protected PAPBinTable _pbt;

        /** Contains formatting properties for sections.*/
        protected SectionTable _st;

        /** Holds fonts for this document.*/
        protected FontTable _ft;

        /** Hold list tables */
        protected ListTables _lt;

        /** main document stream buffer*/
        protected byte[] _mainStream;

        protected HWPFDocumentCore()
            : base(null, null)
        {

        }

        /**
         * Takens an InputStream, verifies that it's not RTF, builds a
         *  POIFSFileSystem from it, and returns that.
         */
        public static POIFSFileSystem VerifyAndBuildPOIFS(Stream istream)
        {
            // Open a PushbackInputStream, so we can peek at the first few bytes
            PushbackStream pis = new PushbackStream(istream);
            byte[] first6 = new byte[6];
            pis.Read(first6, 0, 6);

            // Does it start with {\rtf ? If so, it's really RTF
            if (first6[0] == '{' && first6[1] == '\\' && first6[2] == 'r'
                && first6[3] == 't' && first6[4] == 'f')
            {
                throw new ArgumentException("The document is really a RTF file");
            }

            // OK, so it's not RTF
            // Open a POIFSFileSystem on the (pushed back) stream
            pis.Unread(6);
            return new POIFSFileSystem(istream);
        }

        /// <summary>
        /// This constructor loads a Word document from an stream.
        /// </summary>
        /// <param name="istream">The InputStream that Contains the Word document.</param>
        public HWPFDocumentCore(Stream istream)
            : this(VerifyAndBuildPOIFS(istream))
        {
            //do Ole stuff

        }
 
        /// <summary>
        /// This constructor loads a Word document from a POIFSFileSystem
        /// </summary>
        /// <param name="pfilesystem">The POIFSFileSystem that Contains the Word document.</param>
        public HWPFDocumentCore(POIFSFileSystem pfilesystem)
            : this(pfilesystem.Root)
        {

        }
        /// <summary>
        /// This constructor loads a Word document from a specific point in a POIFSFileSystem, probably not the default.Used typically to open embeded documents.
        /// </summary>
        /// <param name="directory">The POIFSFileSystem that Contains the Word document.</param>
        /// <param name="pfilesystem">If there is an unexpected IOException from the passed in POIFSFileSystem.</param>
        public HWPFDocumentCore(DirectoryNode directory)
            : base(directory)
        {
            // Sort out the hpsf properties


            // read in the main stream.
            DocumentEntry documentProps = (DocumentEntry)
               directory.GetEntry(STREAM_WORD_DOCUMENT);
            _mainStream = new byte[documentProps.Size];

            directory.CreatePOIFSDocumentReader(STREAM_WORD_DOCUMENT).Read(_mainStream);

            // Create our FIB, and check for the doc being encrypted
            _fib = new FileInformationBlock(_mainStream);
            if (_fib.IsFEncrypted())
            {
                throw new EncryptedDocumentException("Cannot process encrypted word files!");
            }

            DirectoryEntry objectPoolEntry;
            try
            {
                objectPoolEntry = (DirectoryEntry)directory.GetEntry(STREAM_OBJECT_POOL);
            }
            catch (FileNotFoundException exc)
            {
                objectPoolEntry = null;
            }
            _objectPool = new ObjectPoolImpl(objectPoolEntry);
        }

        /// <summary>
        /// the range which covers the whole of the document, but excludes any headers and footers.
        /// </summary>
        public abstract Range GetRange();

        /// <summary>
        /// the range that covers all text in the file, including main text,footnotes, headers and comments
        /// </summary>
        public abstract Range GetOverallRange();

            /**
     * Returns document text, i.e. text information from all text pieces,
     * including OLE descriptions and field codes
     */
    public String GetDocumentText() {
        return Text.ToString();
    }

    /**
     * Internal method to access document text
     */
    
        public abstract StringBuilder Text{get;}

        public abstract TextPieceTable TextTable { get; }


        public CHPBinTable CharacterTable
        {
            get
            {
                return _cbt;
            }
        }

        public PAPBinTable ParagraphTable
        {
            get
            {
                return _pbt;
            }
        }

        public SectionTable SectionTable
        {
            get
            {
                return _st;
            }
        }

        public StyleSheet GetStyleSheet()
        {
            return _ss;
        }

        public ListTables GetListTables()
        {
            return _lt;
        }

        public FontTable GetFontTable()
        {
            return _ft;
        }

        public FileInformationBlock GetFileInformationBlock()
        {
            return _fib;
        }

        public ObjectsPool GetObjectsPool()
        {
            return _objectPool;
        }
    }
}