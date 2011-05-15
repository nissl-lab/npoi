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

    using NPOI.HWPF.Model.IO;
    using NPOI.HWPF.Model;
    using NPOI.HWPF.UserModel;
    using System.IO;
    using System;
    using NPOI.POIFS.FileSystem;
    using NPOI.POIFS.Common;
    using System.Collections.Generic;


    /**
     *
     * This class acts as the bucket that we throw all of the Word data structures
     * into.
     *
     * @author Ryan Ackley
     */
    public class HWPFDocument : HWPFDocumentCore
    {
        /** And for making sense of CP lengths in the FIB */
        protected CPSplitCalculator _cpSplit;

        /** table stream buffer*/
        protected byte[] _tableStream;

        /** data stream buffer*/
        protected byte[] _dataStream;

        /** Document wide Properties*/
        protected DocumentProperties _dop;

        /** Contains text of the document wrapped in a obfuscated Word data
        * structure*/
        protected ComplexFileTable _cft;

        protected TextPieceTable _tpt;

        /** Holds the save history for this document. */
        protected SavedByTable _sbt;

        /** Holds the revision mark authors for this document. */
        protected RevisionMarkAuthorTable _rmat;

        /** Holds pictures table */
        protected PicturesTable _pictures;

        /** Holds FSBA (shape) information */
        protected FSPATable _fspa;

        /** Escher Drawing Group information */
        protected EscherRecordHolder _dgg;

        /** Holds Office Art objects */
        protected ShapesTable _officeArts;

        protected HWPFDocument()
            : base()
        {

        }

        /**
         * This constructor loads a Word document from an InputStream.
         *
         * @param istream The InputStream that Contains the Word document.
         * @throws IOException If there is an unexpected IOException from the passed
         *         in InputStream.
         */
        public HWPFDocument(Stream istream)
            : this(VerifyAndBuildPOIFS(istream))
        {
            //do Ole stuff

        }

        /**
         * This constructor loads a Word document from a POIFSFileSystem
         *
         * @param pfilesystem The POIFSFileSystem that Contains the Word document.
         * @throws IOException If there is an unexpected IOException from the passed
         *         in POIFSFileSystem.
         */
        public HWPFDocument(POIFSFileSystem pfilesystem)
            : this(pfilesystem.Root, pfilesystem)
        {

        }

        /**
         * This constructor loads a Word document from a specific point
         *  in a POIFSFileSystem, probably not the default.
         * Used typically to open embeded documents.
         *
         * @param pfilesystem The POIFSFileSystem that Contains the Word document.
         * @throws IOException If there is an unexpected IOException from the passed
         *         in POIFSFileSystem.
         */
        public HWPFDocument(DirectoryNode directory, POIFSFileSystem pfilesystem)
            : base(directory, pfilesystem)
        {
            // Load the main stream and FIB
            // Also handles HPSF bits

            // Do the CP Split
            _cpSplit = new CPSplitCalculator(_fib);

            // Is this document too old for us?
            if (_fib.GetNFib() < 106)
            {
                throw new OldWordFileFormatException("The document is too old - Word 95 or older. Try HWPFOldDocument instead?");
            }

            // use the fib to determine the name of the table stream.
            String name = "0Table";
            if (_fib.IsFWhichTblStm())
            {
                name = "1Table";
            }

            // Grab the table stream.
            DocumentEntry tableProps;
            try
            {
                tableProps =
                    (DocumentEntry)directory.GetEntry(name);
            }
            catch (FileNotFoundException)
            {
                throw new InvalidOperationException("Table Stream '" + name + "' wasn't found - Either the document is corrupt, or is Word95 (or earlier)");
            }

            // read in the table stream.
            _tableStream = new byte[tableProps.Size];
            directory.CreatePOIFSDocumentReader(name).Read(_tableStream);

            _fib.FillVariableFields(_mainStream, _tableStream);

            // read in the data stream.
            try
            {
                DocumentEntry dataProps =
                    (DocumentEntry)directory.GetEntry("Data");
                _dataStream = new byte[dataProps.Size];
                filesystem.CreatePOIFSDocumentReader("Data").Read(_dataStream);
            }
            catch (FileNotFoundException)
            {
                _dataStream = new byte[0];
            }

            // Get the cp of the start of text in the main stream
            // The latest spec doc says this is always zero!
            int fcMin = 0;
            //fcMin = _fib.GetFcMin()

            // Start to load up our standard structures.
            _dop = new DocumentProperties(_tableStream, _fib.GetFcDop());
            _cft = new ComplexFileTable(_mainStream, _tableStream, _fib.GetFcClx(), fcMin);
            _tpt = _cft.GetTextPieceTable();

            // Word XP and later all put in a zero Filled buffer in
            //  front of the text. This screws up the system for offSets,
            //  which assume we always start at zero. This is an adjustment.
            int cpMin = _tpt.CpMin;

            // Now load the rest of the properties, which need to be adjusted
            //  for where text really begin
            _cbt = new CHPBinTable(_mainStream, _tableStream, _fib.GetFcPlcfbteChpx(), _fib.GetLcbPlcfbteChpx(), cpMin, _tpt);
            _pbt = new PAPBinTable(_mainStream, _tableStream, _dataStream, _fib.GetFcPlcfbtePapx(), _fib.GetLcbPlcfbtePapx(), cpMin, _tpt);

            // Read FSPA and Escher information
            _fspa = new FSPATable(_tableStream, _fib.GetFcPlcspaMom(), _fib.GetLcbPlcspaMom(), TextTable.TextPieces);

            if (_fib.GetFcDggInfo() != 0)
            {
                _dgg = new EscherRecordHolder(_tableStream, _fib.GetFcDggInfo(), _fib.GetLcbDggInfo());
            }
            else
            {
                _dgg = new EscherRecordHolder();
            }

            // read in the pictures stream
            _pictures = new PicturesTable(this, _dataStream, _mainStream, _fspa, _dgg);
            // And the art shapes stream
            _officeArts = new ShapesTable(_tableStream, _fib);

            _st = new SectionTable(_mainStream, _tableStream, _fib.GetFcPlcfsed(), _fib.GetLcbPlcfsed(), fcMin, _tpt, _cpSplit);
            _ss = new StyleSheet(_tableStream, _fib.GetFcStshf());
            _ft = new FontTable(_tableStream, _fib.GetFcSttbfffn(), _fib.GetLcbSttbfffn());

            int listOffset = _fib.GetFcPlcfLst();
            int lfoOffset = _fib.GetFcPlfLfo();
            if (listOffset != 0 && _fib.GetLcbPlcfLst() != 0)
            {
                _lt = new ListTables(_tableStream, _fib.GetFcPlcfLst(), _fib.GetFcPlfLfo());
            }

            int sbtOffset = _fib.GetFcSttbSavedBy();
            int sbtLength = _fib.GetLcbSttbSavedBy();
            if (sbtOffset != 0 && sbtLength != 0)
            {
                _sbt = new SavedByTable(_tableStream, sbtOffset, sbtLength);
            }

            int rmarkOffset = _fib.GetFcSttbfRMark();
            int rmarkLength = _fib.GetLcbSttbfRMark();
            if (rmarkOffset != 0 && rmarkLength != 0)
            {
                _rmat = new RevisionMarkAuthorTable(_tableStream, rmarkOffset, rmarkLength);
            }

            PlexOfCps plc = new PlexOfCps(_tableStream, _fib.GetFcPlcffldMom(), _fib.GetLcbPlcffldMom(), 2);
            for (int x = 0; x < plc.Length; x++)
            {
                GenericPropertyNode node = plc.GetProperty(x);
                byte[] fld = node.Bytes;
                int breakpoint = 0;
            }
        }

        public override TextPieceTable TextTable
        {
            get
            {
                return _cft.GetTextPieceTable();
            }
        }

        public CPSplitCalculator GetCPSplitCalculator()
        {
            return _cpSplit;
        }

        public DocumentProperties GetDocProperties()
        {
            return _dop;
        }

        /**
         * Returns the range that covers all text in the
         *  file, including main text, footnotes, headers
         *  and comments
         */
        public Range GetOverallRange()
        {
            // hack to get the ending cp of the document, Have to revisit this.
            PropertyNode p = _tpt.TextPieces[_tpt.TextPieces.Count - 1];

            return new Range(0, p.End, this);
        }

        /**
         * Returns the range which covers the whole of the
         *  document, but excludes any headers and footers.
         */
        public override Range GetRange()
        {
            // First up, trigger a full-recalculate
            // Needed in case of deletes etc
            GetOverallRange();

            // Now, return the real one
            return new Range(
                    _cpSplit.GetMainDocumentStart(),
                    _cpSplit.GetMainDocumentEnd(),
                    this
            );
        }

        /**
         * Returns the range which covers all the Footnotes.
         */
        public Range GetFootnoteRange()
        {
            return new Range(
                    _cpSplit.GetFootnoteStart(),
                    _cpSplit.GetFootnoteEnd(),
                    this
            );
        }

        /**
         * Returns the range which covers all the Endnotes.
        */
        public Range GetEndnoteRange()
        {
            return new Range(
                            _cpSplit.GetEndNoteStart(),
                            _cpSplit.GetEndNoteEnd(),
                            this
        );
        }

        /**
         * Returns the range which covers all the Endnotes.
        */
        public Range GetCommentsRange()
        {
            return new Range(
                            _cpSplit.GetCommentsStart(),
                            _cpSplit.GetCommentsEnd(),
                            this
        );
        }

        /**
         * Returns the range which covers all "Header Stories".
         * A header story Contains a header, footer, end note
         *  separators and footnote separators.
         */
        public Range GetHeaderStoryRange()
        {
            return new Range(
                    _cpSplit.GetHeaderStoryStart(),
                    _cpSplit.GetHeaderStoryEnd(),
                    this
            );
        }

        /**
         * Returns the character length of a document.
         * @return the character length of a document
         */
        public int CharacterLength
        {
            get
            {
                List<TextPiece> textPieces = _tpt.TextPieces;
                IEnumerator<TextPiece> textIt = textPieces.GetEnumerator();

                int length = 0;
                while (textIt.MoveNext())
                {
                    TextPiece tp = textIt.Current;
                    length += tp.CharacterLength;
                }
                return length;
            }
        }

        /**
         * Gets a reference to the saved -by table, which holds the save history for the document.
         *
         * @return the saved-by table.
         */
        public SavedByTable GetSavedByTable()
        {
            return _sbt;
        }

        /**
         * Gets a reference to the revision mark author table, which holds the revision mark authors for the document.
         *
         * @return the saved-by table.
         */
        public RevisionMarkAuthorTable GetRevisionMarkAuthorTable()
        {
            return _rmat;
        }

        /**
         * @return PicturesTable object, that is able to extract images from this document
         */
        public PicturesTable GetPicturesTable()
        {
            return _pictures;
        }

        /**
         * @return ShapesTable object, that is able to extract office are shapes from this document
         */
        public ShapesTable GetShapesTable()
        {
            return _officeArts;
        }

        /**
         * Writes out the word file that is represented by an instance of this class.
         *
         * @param out The OutputStream to write to.
         * @throws IOException If there is an unexpected IOException from the passed
         *         in OutputStream.
         */
        public override void Write(Stream out1)
        {
            // Initialize our streams for writing.
            HWPFFileSystem docSys = new HWPFFileSystem();
            HWPFStream mainStream = docSys.GetStream("WordDocument");
            HWPFStream tableStream = docSys.GetStream("1Table");
            //HWPFOutputStream dataStream = docSys.GetStream("Data");
            int tableOffset = 0;

            // FileInformationBlock fib = (FileInformationBlock)_fib.Clone();
            // clear the offSets and sizes in our FileInformationBlock.
            _fib.ClearOffsetsSizes();

            // determine the FileInformationBLock size
            int fibSize = _fib.GetSize();
            fibSize += POIFSConstants.SMALLER_BIG_BLOCK_SIZE -
                (fibSize % POIFSConstants.SMALLER_BIG_BLOCK_SIZE);

            // preserve space for the FileInformationBlock because we will be writing
            // it after we write everything else.
            byte[] placeHolder = new byte[fibSize];
            mainStream.Write(placeHolder);
            int mainOffset = mainStream.Offset;

            // write out the StyleSheet.
            _fib.SetFcStshf(tableOffset);
            _ss.WriteTo(tableStream);
            _fib.SetLcbStshf(tableStream.Offset - tableOffset);
            tableOffset = tableStream.Offset;

            // get fcMin and fcMac because we will be writing the actual text with the
            // complex table.
            int fcMin = mainOffset;

            // write out the Complex table, includes text.
            _fib.SetFcClx(tableOffset);
            _cft.WriteTo(docSys);
            _fib.SetLcbClx(tableStream.Offset - tableOffset);
            tableOffset = tableStream.Offset;
            int fcMac = mainStream.Offset;

            // write out the CHPBinTable.
            _fib.SetFcPlcfbteChpx(tableOffset);
            _cbt.WriteTo(docSys, fcMin);
            _fib.SetLcbPlcfbteChpx(tableStream.Offset - tableOffset);
            tableOffset = tableStream.Offset;

            // write out the PAPBinTable.
            _fib.SetFcPlcfbtePapx(tableOffset);
            _pbt.WriteTo(docSys, fcMin);
            _fib.SetLcbPlcfbtePapx(tableStream.Offset - tableOffset);
            tableOffset = tableStream.Offset;

            // write out the SectionTable.
            _fib.SetFcPlcfsed(tableOffset);
            _st.WriteTo(docSys, fcMin);
            _fib.SetLcbPlcfsed(tableStream.Offset - tableOffset);
            tableOffset = tableStream.Offset;

            // write out the list tables
            if (_lt != null)
            {
                _fib.SetFcPlcfLst(tableOffset);
                _lt.WriteListDataTo(tableStream);
                _fib.SetLcbPlcfLst(tableStream.Offset - tableOffset);

                _fib.SetFcPlfLfo(tableStream.Offset);
                _lt.WriteListOverridesTo(tableStream);
                _fib.SetLcbPlfLfo(tableStream.Offset - tableOffset);
                tableOffset = tableStream.Offset;
            }

            // write out the saved-by table.
            if (_sbt != null)
            {
                _fib.SetFcSttbSavedBy(tableOffset);
                _sbt.WriteTo(tableStream);
                _fib.SetLcbSttbSavedBy(tableStream.Offset - tableOffset);

                tableOffset = tableStream.Offset;
            }

            // write out the revision mark authors table.
            if (_rmat != null)
            {
                _fib.SetFcSttbfRMark(tableOffset);
                _rmat.WriteTo(tableStream);
                _fib.SetLcbSttbfRMark(tableStream.Offset - tableOffset);

                tableOffset = tableStream.Offset;
            }

            // write out the FontTable.
            _fib.SetFcSttbfffn(tableOffset);
            _ft.WriteTo(docSys);
            _fib.SetLcbSttbfffn(tableStream.Offset - tableOffset);
            tableOffset = tableStream.Offset;

            // write out the DocumentProperties.
            _fib.SetFcDop(tableOffset);
            byte[] buf = new byte[_dop.GetSize()];
            _fib.SetLcbDop(_dop.GetSize());
            _dop.Serialize(buf, 0);
            tableStream.Write(buf);

            // set some variables in the FileInformationBlock.
            _fib.SetFcMin(fcMin);
            _fib.SetFcMac(fcMac);
            _fib.SetCbMac(mainStream.Offset);

            // make sure that the table, doc and data streams use big blocks.
            byte[] mainBuf = mainStream.ToArray();
            if (mainBuf.Length < 4096)
            {
                byte[] tempBuf = new byte[4096];
                Array.Copy(mainBuf, 0, tempBuf, 0, mainBuf.Length);
                mainBuf = tempBuf;
            }

            // write out the FileInformationBlock.
            //_fib.Serialize(mainBuf, 0);
            _fib.WriteTo(mainBuf, tableStream);

            byte[] tableBuf = tableStream.ToArray();
            if (tableBuf.Length < 4096)
            {
                byte[] tempBuf = new byte[4096];
                Array.Copy(tableBuf, 0, tempBuf, 0, tableBuf.Length);
                tableBuf = tempBuf;
            }

            byte[] dataBuf = _dataStream;
            if (dataBuf == null)
            {
                dataBuf = new byte[4096];
            }
            if (dataBuf.Length < 4096)
            {
                byte[] tempBuf = new byte[4096];
                Array.Copy(dataBuf, 0, tempBuf, 0, dataBuf.Length);
                dataBuf = tempBuf;
            }


            // spit out the Word document.
            POIFSFileSystem pfs = new POIFSFileSystem();
            pfs.CreateDocument(new MemoryStream(mainBuf), "WordDocument");
            pfs.CreateDocument(new MemoryStream(tableBuf), "1Table");
            pfs.CreateDocument(new MemoryStream(dataBuf), "Data");
            WriteProperties(pfs);

            pfs.WriteFileSystem(out1);
        }

        public byte[] GetDataStream()
        {
            return _dataStream;
        }
        public byte[] GetTableStream()
        {
            return _tableStream;
        }

        public int RegisterList(HWPFList list)
        {
            if (_lt == null)
            {
                _lt = new ListTables();
            }
            return _lt.AddList(list.GetListData(), list.GetOverride());
        }

        public void Delete(int start, int length)
        {
            Range r = new Range(start, start + length, this);
            r.Delete();
        }
    }

}