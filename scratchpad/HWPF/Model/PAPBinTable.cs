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

using System.Collections.Generic;
using NPOI.POIFS.Common;
using NPOI.Util;
using NPOI.HWPF.SPRM;
using NPOI.HWPF.Model.IO;
using System.IO;
using System;
using System.Text;
namespace NPOI.HWPF.Model
{


    /**
     * This class represents the bin table of Word document but it also serves as a
     * holder for all of the paragraphs of document that have been loaded into
     * memory.
     *
     * @author Ryan Ackley
     */
    public class PAPBinTable
    {
        protected List<PAPX> _paragraphs = new List<PAPX>();
        byte[] _dataStream;

        /** So we can know if things are unicode or not */
        private TextPieceTable tpt;

        public PAPBinTable()
        { 
        }

        [Obsolete]
        public PAPBinTable(byte[] documentStream, byte[] tableStream,
                byte[] dataStream, int offset, int size, int fcMin,
                TextPieceTable tpt) :
            this(documentStream, tableStream, dataStream, offset, size, tpt)
        {

        }

        public PAPBinTable(byte[] documentStream, byte[] tableStream, byte[] dataStream, int offset,
                           int size, CharIndexTranslator charIndexTranslator)
        {
            PlexOfCps binTable = new PlexOfCps(tableStream, offset, size, 4);

            int length = binTable.Length;
            for (int x = 0; x < length; x++)
            {
                GenericPropertyNode node = binTable.GetProperty(x);

                int pageNum = LittleEndian.GetInt(node.Bytes);
                int pageOffset = POIFSConstants.SMALLER_BIG_BLOCK_SIZE * pageNum;

                PAPFormattedDiskPage pfkp = new PAPFormattedDiskPage(documentStream,
                  dataStream, pageOffset, charIndexTranslator);

                int fkpSize = pfkp.Size();

                for (int y = 0; y < fkpSize; y++)
                {
                    PAPX papx = pfkp.GetPAPX(y);
                    _paragraphs.Add(papx);
                }
            }
            _dataStream = dataStream;
        }

        public void Insert(int listIndex, int cpStart, SprmBuffer buf)
        {

            PAPX forInsert = new PAPX(0, 0, buf);

            // Ensure character OffSets are really characters
            forInsert.Start = cpStart;
            forInsert.End = cpStart;

            if (listIndex == _paragraphs.Count)
            {
                _paragraphs.Add(forInsert);
            }
            else
            {
                PAPX currentPap = _paragraphs[listIndex];
                if (currentPap != null && currentPap.Start < cpStart)
                {
                    SprmBuffer ClonedBuf = null;
                    ClonedBuf = (SprmBuffer)currentPap.GetSprmBuf().Clone();


                    // Copy the properties of the one before to afterwards
                    // Will go:
                    //  Original, until insert at point
                    //  New one
                    //  Clone of original, on to the old end
                    PAPX clone = new PAPX(0, 0, ClonedBuf);
                    // Again ensure Contains character based OffSets no matter what
                    clone.Start = (cpStart);
                    clone.End = (currentPap.End);

                    currentPap.End = cpStart;

                    _paragraphs.Insert(listIndex + 1, forInsert);
                    _paragraphs.Insert(listIndex + 2, clone);
                }
                else
                {
                    _paragraphs.Insert(listIndex, forInsert);
                }
            }

        }
        internal class PAPXToFileComparer : IComparer<PAPX>
        {
            Dictionary<PAPX, int> list;
            public PAPXToFileComparer(Dictionary<PAPX, int> list)
            {
                this.list = list;
            }
            #region IComparer<CHPX> Members

            public int Compare(PAPX o1, PAPX o2)
            {
                int i1 = list[o1];
                int i2 = list[o2];
                return i1.CompareTo(i2);
            }
            #endregion
        }
        public void Rebuild(StringBuilder docText,
        ComplexFileTable complexFileTable)
        {
            long start = DateTime.Now.Ticks;

            if (complexFileTable != null)
            {
                SprmBuffer[] sprmBuffers = complexFileTable.GetGrpprls();

                // adding PAPX from fast-saved SPRMs
                foreach (TextPiece textPiece in complexFileTable.GetTextPieceTable()
                        .TextPieces)
                {
                    PropertyModifier prm = textPiece.PieceDescriptor.Prm;
                    if (!prm.IsComplex())
                        continue;
                    int igrpprl = prm.GetIgrpprl();

                    if (igrpprl < 0 || igrpprl >= sprmBuffers.Length)
                    {
                        logger.Log(POILogger.WARN, textPiece
                                + "'s PRM references to unknown grpprl");
                        continue;
                    }

                    bool hasPap = false;
                    SprmBuffer sprmBuffer = sprmBuffers[igrpprl];
                    for (SprmIterator iterator = sprmBuffer.Iterator(); iterator
                            .HasNext(); )
                    {
                        SprmOperation sprmOperation = iterator.Next();
                        if (sprmOperation.Type == SprmOperation.TYPE_PAP)
                        {
                            hasPap = true;
                            break;
                        }
                    }

                    if (hasPap)
                    {
                        SprmBuffer newSprmBuffer = new SprmBuffer(2);
                        newSprmBuffer.Append(sprmBuffer.ToByteArray());

                        PAPX papx = new PAPX(textPiece.Start,
                                textPiece.End, newSprmBuffer);
                        _paragraphs.Add(papx);
                    }
                }

                logger.Log(POILogger.DEBUG,
                        "Merged (?) with PAPX from complex file table in ",
                        DateTime.Now.Ticks - start,
                        " ms (", _paragraphs.Count,
                        " elements in total)");
                start = DateTime.Now.Ticks;
            }

            List<PAPX> oldPapxSortedByEndPos = new List<PAPX>(_paragraphs);
            oldPapxSortedByEndPos.Sort(
                    (IComparer<PAPX>)PropertyNode.PAPXComparator.instance);

            logger.Log(POILogger.DEBUG, "PAPX sorted by end position in ",
                    DateTime.Now.Ticks - start, " ms");
            start = DateTime.Now.Ticks;

            Dictionary<PAPX, int> papxToFileOrder = new Dictionary<PAPX, int>();
            int counter = 0;
            foreach (PAPX papx in _paragraphs)
            {
                papxToFileOrder[papx] = counter++;
            }

            logger.Log(POILogger.DEBUG, "PAPX's order map created in ",
                    DateTime.Now.Ticks - start, " ms");
            start = DateTime.Now.Ticks;

            List<PAPX> newPapxs = new List<PAPX>();
            int lastParStart = 0;
            int lastPapxIndex = 0;
            for (int charIndex = 0; charIndex < docText.Length; charIndex++)
            {
                char c = docText[charIndex];
                if (c != 13 && c != 7 && c != 12)
                    continue;

                int startInclusive = lastParStart;
                int endExclusive = charIndex + 1;

                bool broken = false;
                List<PAPX> papxs = new List<PAPX>();
                for (int papxIndex = lastPapxIndex; papxIndex < oldPapxSortedByEndPos
                        .Count; papxIndex++)
                {
                    broken = false;
                    PAPX papx = oldPapxSortedByEndPos[papxIndex];


                    if (papx.End - 1 > charIndex)
                    {
                        lastPapxIndex = papxIndex;
                        broken = true;
                        break;
                    }

                    papxs.Add(papx);
                }
                if (!broken)
                {
                    lastPapxIndex = oldPapxSortedByEndPos.Count - 1;
                }

                if (papxs.Count == 0)
                {
                    logger.Log(POILogger.WARN, "Paragraph [",
                             startInclusive, "; ",
                             endExclusive,
                            ") has no PAPX. Creating new one.");
                    // create it manually
                    PAPX papx = new PAPX(startInclusive, endExclusive,
                            new SprmBuffer(2));
                    newPapxs.Add(papx);

                    lastParStart = endExclusive;
                    continue;
                }

                if (papxs.Count == 1)
                {
                    // can we reuse existing?
                    PAPX existing = papxs[0];
                    if (existing.Start == startInclusive
                            && existing.End == endExclusive)
                    {
                        newPapxs.Add(existing);
                        lastParStart = endExclusive;
                        continue;
                    }
                }
                PAPXToFileComparer papxFileOrderComparator = new PAPXToFileComparer(papxToFileOrder);
                // restore file order of PAPX
                papxs.Sort(papxFileOrderComparator);

                SprmBuffer sprmBuffer = null;
                foreach (PAPX papx in papxs)
                {
                    if (sprmBuffer == null)
                        sprmBuffer = (SprmBuffer)papx.GetSprmBuf().Clone();

                    else
                        sprmBuffer.Append(papx.GetGrpprl(), 2);
                }
                PAPX newPapx = new PAPX(startInclusive, endExclusive, sprmBuffer);
                newPapxs.Add(newPapx);

                lastParStart = endExclusive;
                continue;
            }
            this._paragraphs = new List<PAPX>(newPapxs);

            logger.Log(POILogger.DEBUG, "PAPX rebuilded from document text in ",
                    DateTime.Now.Ticks - start, " ms (",
                     _paragraphs.Count, " elements)");
            start = DateTime.Now.Ticks;
        }

        private static POILogger logger = POILogFactory
        .GetLogger(typeof(PAPBinTable));

        public void AdjustForDelete(int listIndex, int offset, int Length)
        {
            int size = _paragraphs.Count;
            int endMark = offset + Length;
            int endIndex = listIndex;

            PAPX papx = _paragraphs[endIndex];
            while (papx.End < endMark)
            {
                papx = _paragraphs[++endIndex];
            }
            if (listIndex == endIndex)
            {
                papx = _paragraphs[endIndex];
                papx.End = ((papx.End - endMark) + offset);
            }
            else
            {
                papx = _paragraphs[listIndex];
                papx.End = (offset);
                for (int x = listIndex + 1; x < endIndex; x++)
                {
                    papx = _paragraphs[x];
                    papx.Start = (offset);
                    papx.End = (offset);
                }
                papx = _paragraphs[endIndex];
                papx.End = ((papx.End - endMark) + offset);
            }

            for (int x = endIndex + 1; x < size; x++)
            {
                papx = _paragraphs[x];
                papx.Start = (papx.Start - Length);
                papx.End = (papx.End - Length);
            }
        }


        public void AdjustForInsert(int listIndex, int Length)
        {
            int size = _paragraphs.Count;
            PAPX papx = (PAPX)_paragraphs[listIndex];
            papx.End = (papx.End + Length);

            for (int x = listIndex + 1; x < size; x++)
            {
                papx = (PAPX)_paragraphs[x];
                papx.Start = (papx.Start + Length);
                papx.End = (papx.End + Length);
            }
        }


        public List<PAPX> GetParagraphs()
        {
            return _paragraphs;
        }

        [Obsolete]
        public void WriteTo(HWPFFileSystem sys, CharIndexTranslator translator)
        {
            HWPFStream wordDocumentStream = sys.GetStream("WordDocument");
            HWPFStream tableStream = sys.GetStream("1Table");

            WriteTo(wordDocumentStream, tableStream, translator);
        }
        public void WriteTo(HWPFStream docStream,
            HWPFStream tableStream, CharIndexTranslator translator )

        {
            PlexOfCps binTable = new PlexOfCps(4);

            // each FKP must start on a 512 byte page.
            int docOffset = docStream.Offset;
            int mod = docOffset % POIFSConstants.SMALLER_BIG_BLOCK_SIZE;
            if (mod != 0)
            {
                byte[] pAdding = new byte[POIFSConstants.SMALLER_BIG_BLOCK_SIZE - mod];
                docStream.Write(pAdding);
            }

            // get the page number for the first fkp
            docOffset = docStream.Offset;
            int pageNum = docOffset / POIFSConstants.SMALLER_BIG_BLOCK_SIZE;

            // get the ending fc
            int endingFc = ((PropertyNode)_paragraphs[_paragraphs.Count - 1]).End;


            List<PAPX> overflow = _paragraphs;
            do
            {
                PropertyNode startingProp = (PropertyNode)overflow[0];
                int start = translator.GetByteIndex(startingProp.Start);

                PAPFormattedDiskPage pfkp = new PAPFormattedDiskPage(_dataStream);
                pfkp.Fill(overflow);

                byte[] bufFkp = pfkp.ToByteArray(tableStream,translator);
                docStream.Write(bufFkp);
                overflow = pfkp.GetOverflow();

                int end = endingFc;
                if (overflow != null)
                {
                    end = translator.GetByteIndex(overflow[0].Start);
                }

                byte[] intHolder = new byte[4];
                LittleEndian.PutInt(intHolder, pageNum++);
                binTable.AddProperty(new GenericPropertyNode(start, end, intHolder));

            }
            while (overflow != null);
            byte[] bytes = binTable.ToByteArray();
            tableStream.Write(bytes, 0, bytes.Length);
        }


    }
}
