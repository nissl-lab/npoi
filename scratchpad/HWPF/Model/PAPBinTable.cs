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

        public PAPBinTable(byte[] documentStream, byte[] tableStream, byte[] dataStream, int offset,
                           int size, int fcMin, TextPieceTable tpt)
        {
            PlexOfCps binTable = new PlexOfCps(tableStream, offset, size, 4);
            this.tpt = tpt;

            int length = binTable.Length;
            for (int x = 0; x < length; x++)
            {
                GenericPropertyNode node = binTable.GetProperty(x);

                int pageNum = LittleEndian.GetInt(node.Bytes);
                int pageOffset = POIFSConstants.SMALLER_BIG_BLOCK_SIZE * pageNum;

                PAPFormattedDiskPage pfkp = new PAPFormattedDiskPage(documentStream,
                  dataStream, pageOffset, fcMin, tpt);

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

            PAPX forInsert = new PAPX(0, 0, tpt, buf, _dataStream);

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
                    PAPX clone = new PAPX(0, 0, tpt, ClonedBuf, _dataStream);
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
            int endingFc = ((PropertyNode)_paragraphs[_paragraphs.Count - 1]).End;
            endingFc += fcMin;


            List<PAPX> overflow = _paragraphs;
            do
            {
                PropertyNode startingProp = (PropertyNode)overflow[0];
                int start = startingProp.Start + fcMin;

                PAPFormattedDiskPage pfkp = new PAPFormattedDiskPage(_dataStream);
                pfkp.Fill(overflow);

                byte[] bufFkp = pfkp.ToArray(fcMin);
                docStream.Write(bufFkp);
                overflow = pfkp.GetOverflow();

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
            tableStream.Write(bytes,0, bytes.Length);
        }


    }
}


