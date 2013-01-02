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

namespace NPOI.HWPF.Model
{

    using NPOI.Util;
    using NPOI.POIFS.Common;

    /**
     * This class holds all of the character formatting 
     *  properties from Old (Word 6 / Word 95) documents.
     * Unlike with Word 97+, it all Gets held in the
     *  same stream.
     * In common with the rest of the old support, it 
     *  is read only
     */
    public class OldCHPBinTable : CHPBinTable
    {
        /**
         * Constructor used to read an old-style binTable
         *  in from a Word document.
         *
         * @param documentStream
         * @param offset
         * @param size
         * @param fcMin
         */
        public OldCHPBinTable(byte[] documentStream, int OffSet,
                           int size, int fcMin, TextPieceTable tpt)
        {
            PlexOfCps binTable = new PlexOfCps(documentStream, OffSet, size, 2);

            int length = binTable.Length;
            for (int x = 0; x < length; x++)
            {
                GenericPropertyNode node = binTable.GetProperty(x);

                int pageNum = LittleEndian.GetShort(node.Bytes);
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
    }
}

