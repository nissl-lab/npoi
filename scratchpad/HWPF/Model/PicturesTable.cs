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

using NPOI.Util;
using NPOI.DDF;
using System.Collections.Generic;
using NPOI.HWPF.UserModel;
using System.Collections;
using System;
namespace NPOI.HWPF.Model
{

    /**
     * Holds information about all pictures embedded in Word Document either via "Insert -> Picture -> From File" or via
     * clipboard. Responsible for images extraction and determining whether some document's piece Contains embedded image.
     * Analyzes raw data bytestream 'Data' (where Word stores all embedded objects) provided by HWPFDocument.
     *
     * Word stores images as is within so called "Data stream" - the stream within a Word docfile Containing various data
     * that hang off of characters in the main stream. For example, binary data describing in-line pictures and/or
     * formfields an also embedded objects-native data. Word picture structures are concatenated one after the other in
     * the data stream if the document Contains pictures.
     * Data stream is easily reachable via HWPFDocument._dataStream property.
     * A picture is represented in the document text stream as a special character, an Unicode \u0001 whose
     * CharacterRun.IsSpecial() returns true. The file location of the picture in the Word binary file is accessed
     * via CharacterRun.GetPicOffSet(). The CharacterRun.GetPicOffSet() is a byte offset into the data stream.
     * Beginning at the position recorded in picOffSet, a header data structure, will be stored.
     *
     * @author Dmitry Romanov
     */
    public class PicturesTable
    {
            private static POILogger logger = POILogFactory
            .GetLogger( typeof(PicturesTable) );

        static int TYPE_IMAGE = 0x08;
        static int TYPE_IMAGE_WORD2000 = 0x00;
        static int TYPE_IMAGE_PASTED_FROM_CLIPBOARD = 0xA;
        static int TYPE_IMAGE_PASTED_FROM_CLIPBOARD_WORD2000 = 0x2;
        static int TYPE_HORIZONTAL_LINE = 0xE;
        static int BLOCK_TYPE_OFFSET = 0xE;
        static int MM_MODE_TYPE_OFFSET = 0x6;

        private HWPFDocument _document;
        private byte[] _dataStream;
        private byte[] _mainStream;
        private FSPATable _fspa;
        private EscherRecordHolder _dgg;

        /** @link dependency
         * @stereotype instantiate*/
        /*# Picture lnkPicture; */

        /**
         *
         * @param _document
         * @param _dataStream
         */
        public PicturesTable(HWPFDocument _document, byte[] _dataStream, byte[] _mainStream, FSPATable fspa, EscherRecordHolder dgg)
        {
            this._document = _document;
            this._dataStream = _dataStream;
            this._mainStream = _mainStream;
            this._fspa = fspa;
            this._dgg = dgg;
        }

        /**
         * determines whether specified CharacterRun Contains reference to a picture
         * @param run
         */
        public bool HasPicture(CharacterRun run)
        {
            if (run.IsSpecialCharacter() && !run.IsObj() && !run.IsOle2() && !run.IsData())
            {
                // Image should be in it's own run, or in a run with the end-of-special marker
                if ("\u0001".Equals(run.Text) || "\u0001\u0015".Equals(run.Text))
                {
                    return IsBlockContainsImage(run.GetPicOffset());
                }
            }
            return false;
        }

        public bool HasEscherPicture(CharacterRun run)
        {
            if (run.IsSpecialCharacter() && !run.IsObj() && !run.IsOle2() && !run.IsData() && run.Text.StartsWith("\u0008"))
            {
                return true;
            }
            return false;
        }

        /**
         * determines whether specified CharacterRun Contains reference to a picture
         * @param run
        */
        public bool HasHorizontalLine(CharacterRun run)
        {
            if (run.IsSpecialCharacter() && "\u0001".Equals(run.Text))
            {
                return IsBlockContainsHorizontalLine(run.GetPicOffset());
            }
            return false;
        }

        private bool IsPictureRecognized(short blockType, short mappingModeOfMETAFILEPICT)
        {
            return (blockType == TYPE_IMAGE || blockType == TYPE_IMAGE_PASTED_FROM_CLIPBOARD || (blockType == TYPE_IMAGE_WORD2000 && mappingModeOfMETAFILEPICT == 0x64) || (blockType == TYPE_IMAGE_PASTED_FROM_CLIPBOARD_WORD2000 && mappingModeOfMETAFILEPICT == 0x64));
        }

        private static short GetBlockType(byte[] dataStream, int pictOffset)
        {
            return LittleEndian.GetShort(dataStream, pictOffset + BLOCK_TYPE_OFFSET);
        }

        private static short GetMmMode(byte[] dataStream, int pictOffset)
        {
            return LittleEndian.GetShort(dataStream, pictOffset + MM_MODE_TYPE_OFFSET);
        }

        /**
         * Returns picture object tied to specified CharacterRun
         * @param run
         * @param FillBytes if true, Picture will be returned with Filled byte array that represent picture's contents. If you don't want
         * to have that byte array in memory but only write picture's contents to stream, pass false and then use Picture.WriteImageContent
         * @see Picture#WriteImageContent(java.io.OutputStream)
         * @return a Picture object if picture exists for specified CharacterRun, null otherwise. PicturesTable.hasPicture is used to determine this.
         * @see #hasPicture(NPOI.HWPF.usermodel.CharacterRun)
         */
        public Picture ExtractPicture(CharacterRun run, bool FillBytes)
        {
            if (HasPicture(run))
            {
                return new Picture(run.GetPicOffset(), _dataStream, FillBytes);
            }
            return null;
        }

        /**
           * Performs a recursive search for pictures in the given list of escher records.
           *
           * @param escherRecords the escher records.
           * @param pictures the list to populate with the pictures.
           */
        private void SearchForPictures(IList escherRecords, List<Picture> pictures)
        {
            foreach (EscherRecord escherRecord in escherRecords)
            {
                if (escherRecord is EscherBSERecord)
                {
                    EscherBSERecord bse = (EscherBSERecord)escherRecord;
                    EscherBlipRecord blip = bse.BlipRecord;
                    if (blip != null)
                    {
                        pictures.Add(new Picture(blip.PictureData));
                    }
                    else if (bse.Offset > 0)
                    {
                        try
                        {
                            // Blip stored in delay stream, which in a word doc, is the main stream
                            IEscherRecordFactory recordFactory = new DefaultEscherRecordFactory();
                            EscherRecord record = recordFactory.CreateRecord(_mainStream, bse.Offset);

                            if (record is EscherBlipRecord)
                            {
                                record.FillFields(_mainStream, bse.Offset, recordFactory);
                                blip = (EscherBlipRecord)record;
                                pictures.Add(new Picture(blip.PictureData));
                            }

                        }
                        catch (Exception exc)
                        {
                            logger.Log(
                                    POILogger.WARN,
                                    "Unable to load picture from BLIB record at offset #",
                                    bse.Offset, exc);
                        }
                    }
                }

                // Recursive call.
                SearchForPictures(escherRecord.ChildRecords, pictures);
            }
        }

        /**
         * Not all documents have all the images concatenated in the data stream
         * although MS claims so. The best approach is to scan all character Runs.
         *
         * @return a list of Picture objects found in current document
         */
        public List<Picture> GetAllPictures()
        {
            List<Picture> pictures = new List<Picture>();

            Range range = _document.GetOverallRange();
            for (int i = 0; i < range.NumCharacterRuns; i++)
            {
                CharacterRun run = range.GetCharacterRun(i);

                if (run == null)
                {
                    continue;
                }

                Picture picture = ExtractPicture(run, false);
                if (picture != null)
                {
                    pictures.Add(picture);
                }
            }

            SearchForPictures(_dgg.EscherRecords, pictures);

            return pictures;
        }

        private bool IsBlockContainsImage(int i)
        {
            return IsPictureRecognized(GetBlockType(_dataStream, i), GetMmMode(_dataStream, i));
        }

        private bool IsBlockContainsHorizontalLine(int i)
        {
            return GetBlockType(_dataStream, i) == TYPE_HORIZONTAL_LINE && GetMmMode(_dataStream, i) == 0x64;
        }

    }
}

