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
using System;
namespace NPOI.HWPF.Model
{


    /**
     * This class holds all of the section formatting 
     *  properties from Old (Word 6 / Word 95) documents.
     * Unlike with Word 97+, it all Gets held in the
     *  same stream.
     * In common with the rest of the old support, it 
     *  is read only
     */
    public class OldSectionTable : SectionTable
    {
        public OldSectionTable(byte[] documentStream, int OffSet,
                            int size, int fcMin,
                            TextPieceTable tpt)
        {
            PlexOfCps sedPlex = new PlexOfCps(documentStream, OffSet, size, 12);
            CharIsBytes charConv = new CharIsBytes(tpt);

            int length = sedPlex.Length;

            for (int x = 0; x < length; x++)
            {
                GenericPropertyNode node = sedPlex.GetProperty(x);
                SectionDescriptor sed = new SectionDescriptor(node.Bytes, 0);

                int fileOffset = sed.GetFc();
                int startAt = node.Start;
                int endAt = node.End;

                // check for the optimization
                if (fileOffset == unchecked((int)0xffffffff))
                {
                    _sections.Add(new SEPX(sed, startAt, endAt, charConv, new byte[0]));
                }
                else
                {
                    // The first short at the offset is the size of the grpprl.
                    int sepxSize = LittleEndian.GetShort(documentStream, fileOffset);
                    // Because we don't properly know about all the details of the old
                    //  section properties, and we're trying to decode them as if they
                    //  were the new ones, we sometimes "need" more data than we have.
                    // As a workaround, have a few extra 0 bytes on the end!
                    byte[] buf = new byte[sepxSize + 2];
                    fileOffset += LittleEndianConstants.SHORT_SIZE;
                    Array.Copy(documentStream, fileOffset, buf, 0, buf.Length);
                    _sections.Add(new SEPX(sed, startAt, endAt, charConv, buf));
                }
            }
        }

        private class CharIsBytes : CharIndexTranslator
        {
            private TextPieceTable tpt;
            internal CharIsBytes(TextPieceTable tpt)
            {
                this.tpt = tpt;
            }

            public int GetCharIndex(int bytePos, int startCP)
            {
                return bytePos;
            }
            public int GetCharIndex(int bytePos)
            {
                return bytePos;
            }

            public bool IsIndexInTable(int bytePos)
            {
                return tpt.IsIndexInTable(bytePos);
            }
            public int LookIndexBackward(int bytePos)
            {
                return tpt.LookIndexBackward(bytePos);
            }
            public int LookIndexForward(int bytePos)
            {
                return tpt.LookIndexForward(bytePos);
            }
        }
    }
}

