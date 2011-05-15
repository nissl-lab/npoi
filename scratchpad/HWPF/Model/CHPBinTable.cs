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

        /** So we can know if things are unicode or not */
        private TextPieceTable tpt;

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
        public CHPBinTable(byte[] documentStream, byte[] tableStream, int offset,
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

                CHPFormattedDiskPage cfkp = new CHPFormattedDiskPage(documentStream,
                  pageOffset, fcMin, tpt);

                int fkpSize = cfkp.Size();

                for (int y = 0; y < fkpSize; y++)
                {
                    _textRuns.Add(cfkp.GetCHPX(y));
                }
            }
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

            CHPX insertChpx = new CHPX(0, 0, tpt, buf);

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
                    CHPX clone = new CHPX(0, 0, tpt, chpx.GetSprmBuf());
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
            tableStream.Write(bytes,0, bytes.Length);
        }
    }
}


