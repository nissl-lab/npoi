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
using NPOI.HWPF.Model.IO;
using System.IO;
using NPOI.HWPF.SPRM;
using System;
namespace NPOI.HWPF.Model
{


    /**
     * This class holds all of the character formatting properties.
     *
     * @author Ryan Ackley
     */
    public class CHPBinTable
    {
        /** List of character properties.*/
        internal List<CHPX> _textRuns = new List<CHPX>();


        public CHPBinTable()
        {
        }

        /**
         * Constructor used to read a binTable in from a Word document.
         *
         * @param documentStream
         * @param tableStream
         * @param offset
         * @param size
         * @param fcMin
         */
        [Obsolete]
        public CHPBinTable(byte[] documentStream, byte[] tableStream, int offset,
                           int size, int fcMin, TextPieceTable tpt):this( documentStream, tableStream, offset, size, tpt )
        {

        }
        /**
 * Constructor used to read a binTable in from a Word document.
 */
        public CHPBinTable(byte[] documentStream, byte[] tableStream, int offset,
                int size, CharIndexTranslator translator)
        {
            PlexOfCps binTable = new PlexOfCps(tableStream, offset, size, 4);

            int length = binTable.Length;
            for (int x = 0; x < length; x++)
            {
                GenericPropertyNode node = binTable.GetProperty(x);

                int pageNum = LittleEndian.GetInt(node.Bytes);
                int pageOffset = POIFSConstants.SMALLER_BIG_BLOCK_SIZE * pageNum;

                CHPFormattedDiskPage cfkp = new CHPFormattedDiskPage(documentStream,
                  pageOffset, translator);

                int fkpSize = cfkp.Size();

                for (int y = 0; y < fkpSize; y++)
                {
                    CHPX chpx = cfkp.GetCHPX(y);
                    if (chpx != null)
                        _textRuns.Add(chpx);
                }
            }
        }

        internal class CHPXToFileComparer : IComparer<CHPX>
        {
            Dictionary<CHPX, int> list;
            public CHPXToFileComparer(Dictionary<CHPX, int> list)
            {
                this.list = list;
            }
            #region IComparer<CHPX> Members

            public int Compare(CHPX o1, CHPX o2)
            {
                int i1 = list[o1];
                int i2 = list[o2];
                return i1.CompareTo(i2);
            }
            #endregion
        }

        private static POILogger logger = POILogFactory
        .GetLogger(typeof(CHPBinTable));

        public void Rebuild(ComplexFileTable complexFileTable)
        {
            long start = DateTime.Now.Ticks;

            if (complexFileTable != null)
            {
                SprmBuffer[] sprmBuffers = complexFileTable.GetGrpprls();

                // adding CHPX from fast-saved SPRMs
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

                    bool hasChp = false;
                    SprmBuffer sprmBuffer = sprmBuffers[igrpprl];
                    for (SprmIterator iterator = sprmBuffer.Iterator(); ; iterator
                            .HasNext())
                    {
                        SprmOperation sprmOperation = iterator.Next();
                        if (sprmOperation.Type == SprmOperation.TYPE_CHP)
                        {
                            hasChp = true;
                            break;
                        }
                    }

                    if (hasChp)
                    {
                        SprmBuffer newSprmBuffer;
                        newSprmBuffer = (SprmBuffer)sprmBuffer.Clone();


                        CHPX chpx = new CHPX(textPiece.Start,
                                textPiece.End, newSprmBuffer);
                        _textRuns.Add(chpx);
                    }
                }
                logger.Log(POILogger.DEBUG,
                        "Merged with CHPX from complex file table in ",
                        DateTime.Now.Ticks - start,
                        " ms (", _textRuns.Count,
                        " elements in total)");
                start = DateTime.Now.Ticks;
            }

            List<CHPX> oldChpxSortedByStartPos = new List<CHPX>(_textRuns);
            oldChpxSortedByStartPos.Sort(
                    (IComparer<CHPX>)PropertyNode.CHPXComparator.instance);

            logger.Log(POILogger.DEBUG, "CHPX sorted by start position in ",
                     DateTime.Now.Ticks - start, " ms");
            start = DateTime.Now.Ticks;

            Dictionary<CHPX, int> chpxToFileOrder = new Dictionary<CHPX, int>();

            int counter = 0;
            foreach (CHPX chpx in _textRuns)
            {
                chpxToFileOrder.Add(chpx, counter++);
            }


            logger.Log(POILogger.DEBUG, "CHPX's order map created in ",
                    DateTime.Now.Ticks - start, " ms");
            start = DateTime.Now.Ticks;

            List<int> textRunsBoundariesList;

            List<int> textRunsBoundariesSet = new List<int>();
            foreach (CHPX chpx in _textRuns)
            {
                textRunsBoundariesSet.Add(chpx.Start);
                textRunsBoundariesSet.Add(chpx.End);
            }
            textRunsBoundariesSet.Remove(0);
            textRunsBoundariesList = new List<int>(
                    textRunsBoundariesSet);
            textRunsBoundariesList.Sort();


            logger.Log(POILogger.DEBUG, "Texts CHPX boundaries collected in ",
                    DateTime.Now.Ticks - start, " ms");
            start = DateTime.Now.Ticks;

            List<CHPX> newChpxs = new List<CHPX>();
            int lastTextRunStart = 0;
            foreach (int objBoundary in textRunsBoundariesList)
            {
                int boundary = objBoundary;

                int startInclusive = lastTextRunStart;
                int endExclusive = boundary;
                lastTextRunStart = endExclusive;

                int startPosition = BinarySearch(oldChpxSortedByStartPos, boundary);
                startPosition = Math.Abs(startPosition);
                while (startPosition >= oldChpxSortedByStartPos.Count)
                    startPosition--;
                while (startPosition > 0
                        && oldChpxSortedByStartPos[startPosition].Start >= boundary)
                    startPosition--;

                List<CHPX> chpxs = new List<CHPX>();
                for (int c = startPosition; c < oldChpxSortedByStartPos.Count; c++)
                {
                    CHPX chpx = oldChpxSortedByStartPos[c];

                    if (boundary < chpx.Start)
                        break;

                    int left = Math.Max(startInclusive, chpx.Start);
                    int right = Math.Min(endExclusive, chpx.End);

                    if (left < right)
                    {
                        chpxs.Add(chpx);
                    }
                }

                if (chpxs.Count == 0)
                {
                    logger.Log(POILogger.WARN, "Text piece [",
                            startInclusive, "; ",
                            endExclusive,
                            ") has no CHPX. Creating new one.");
                    // create it manually
                    CHPX chpx = new CHPX(startInclusive, endExclusive,
                            new SprmBuffer(0));
                    newChpxs.Add(chpx);
                    continue;
                }

                if (chpxs.Count == 1)
                {
                    // can we reuse existing?
                    CHPX existing = chpxs[0];
                    if (existing.Start == startInclusive
                            && existing.End == endExclusive)
                    {
                        newChpxs.Add(existing);
                        continue;
                    }
                }
                CHPXToFileComparer chpxFileOrderComparator = new CHPXToFileComparer(chpxToFileOrder);
                chpxs.Sort(chpxFileOrderComparator);

                SprmBuffer sprmBuffer = new SprmBuffer(0);
                foreach (CHPX chpx in chpxs)
                {
                    sprmBuffer.Append(chpx.GetGrpprl(), 0);
                }
                CHPX newChpx = new CHPX(startInclusive, endExclusive, sprmBuffer);
                newChpxs.Add(newChpx);

                continue;
            }
            this._textRuns = new List<CHPX>(newChpxs);

            logger.Log(POILogger.DEBUG, "CHPX rebuilded in ",
                    DateTime.Now.Ticks - start, " ms (",
                    _textRuns.Count, " elements)");
            start = DateTime.Now.Ticks;

            CHPX previous = null;
            for (List<CHPX>.Enumerator iterator = _textRuns.GetEnumerator(); iterator
                    .MoveNext(); )
            {
                CHPX current = iterator.Current;
                if (previous == null)
                {
                    previous = current;
                    continue;
                }

                if (previous.End == current.Start
                        && Arrays
                                .Equals(previous.GetGrpprl(), current.GetGrpprl()))
                {
                    previous.End = current.End;
                    _textRuns.Remove(current);
                    continue;
                }

                previous = current;
            }

            logger.Log(POILogger.DEBUG, "CHPX compacted in ",
                    DateTime.Now.Ticks - start, " ms (",
                    _textRuns.Count, " elements)");
        }

        private static int BinarySearch(List<CHPX> chpxs, int startPosition)
        {
            int low = 0;
            int high = chpxs.Count - 1;

            while (low <= high)
            {
                int mid = (low + high) >> 1;
                CHPX midVal = chpxs[mid];
                int midValue = midVal.Start;

                if (midValue < startPosition)
                    low = mid + 1;
                else if (midValue > startPosition)
                    high = mid - 1;
                else
                    return mid; // key found
            }
            return -(low + 1); // key not found.
        }

        public void AdjustForDelete(int listIndex, int offset, int Length)
        {
            int size = _textRuns.Count;
            int endMark = offset + Length;
            int endIndex = listIndex;

            CHPX chpx = _textRuns[endIndex];
            while (chpx.End < endMark)
            {
                chpx = _textRuns[++endIndex];
            }
            if (listIndex == endIndex)
            {
                chpx = _textRuns[endIndex];
                chpx.End = ((chpx.End - endMark) + offset);
            }
            else
            {
                chpx = _textRuns[listIndex];
                chpx.End = (offset);
                for (int x = listIndex + 1; x < endIndex; x++)
                {
                    chpx = _textRuns[x];
                    chpx.Start = (offset);
                    chpx.End = (offset);
                }
                chpx = _textRuns[endIndex];
                chpx.End = ((chpx.End - endMark) + offset);
            }

            for (int x = endIndex + 1; x < size; x++)
            {
                chpx = _textRuns[x];
                chpx.Start = (chpx.Start - Length);
                chpx.End = (chpx.End - Length);
            }
        }

        public void Insert(int listIndex, int cpStart, SprmBuffer buf)
        {

            CHPX insertChpx = new CHPX(0, 0, buf);
            
            // Ensure character OffSets are really characters
            insertChpx.Start = (cpStart);
            insertChpx.End = (cpStart);

            if (listIndex == _textRuns.Count)
            {
                _textRuns.Add(insertChpx);
            }
            else
            {
                CHPX chpx = _textRuns[listIndex];
                if (chpx.Start < cpStart)
                {
                    // Copy the properties of the one before to afterwards
                    // Will go:
                    //  Original, until insert at point
                    //  New one
                    //  Clone of original, on to the old end
                    CHPX clone = new CHPX(0, 0, chpx.GetSprmBuf());
                    // Again ensure Contains character based OffSets no matter what
                    clone.Start = (cpStart);
                    clone.End = (chpx.End);

                    chpx.End = (cpStart);

                    _textRuns.Insert(listIndex + 1, insertChpx);
                    _textRuns.Insert(listIndex + 2, clone);
                }
                else
                {
                    _textRuns.Insert(listIndex, insertChpx);
                }
            }
        }

        public void AdjustForInsert(int listIndex, int Length)
        {
            int size = _textRuns.Count;
            CHPX chpx = _textRuns[listIndex];
            chpx.End = (chpx.End + Length);

            for (int x = listIndex + 1; x < size; x++)
            {
                chpx = _textRuns[x];
                chpx.Start = (chpx.Start + Length);
                chpx.End = (chpx.End + Length);
            }
        }

        public List<CHPX> GetTextRuns()
        {
            return _textRuns;
        }

        public void WriteTo(HWPFFileSystem sys, int fcMin)
        {

            HWPFStream docStream = sys.GetStream("WordDocument");
            Stream tableStream = sys.GetStream("1Table");

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
            int endingFc = ((PropertyNode)_textRuns[_textRuns.Count - 1]).End;
            endingFc += fcMin;


            List<CHPX> overflow = _textRuns;
            do
            {
                PropertyNode startingProp = (PropertyNode)overflow[0];
                int start = startingProp.Start + fcMin;

                CHPFormattedDiskPage cfkp = new CHPFormattedDiskPage();
                cfkp.Fill(overflow);

                byte[] bufFkp = cfkp.ToArray(fcMin);
                docStream.Write(bufFkp);
                overflow = cfkp.GetOverflow();

                int end = endingFc;
                if (overflow != null)
                {
                    end = ((PropertyNode)overflow[0]).Start + fcMin;
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