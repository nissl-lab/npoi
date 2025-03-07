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

    using System.Collections.Generic;
    using NPOI.Util;
    using System;
    using NPOI.HWPF.Model.IO;
    /**
     * Represents a PAP FKP. The style properties for paragraph and character Runs
     * are stored in fkps. There are PAP fkps for paragraph properties and CHP fkps
     * for character run properties. The first part of the fkp for both CHP and PAP
     * fkps consists of an array of 4 byte int offsets in the main stream for that
     * Paragraph's or Character Run's text. The ending offset is the next
     * value in the array. For example, if an fkp has X number of Paragraph's
     * stored in it then there are (x + 1) 4 byte ints in the beginning array. The
     * number X is determined by the last byte in a 512 byte fkp.
     *
     * CHP and PAP fkps also store the compressed styles(grpprl) that correspond to
     * the offsets on the front of the fkp. The offset of the grpprls is determined
     * differently for CHP fkps and PAP fkps.
     *
     * @author Ryan Ackley
     */
    public class PAPFormattedDiskPage : FormattedDiskPage
    {

        private static int BX_SIZE = 13;
        private static int FC_SIZE = 4;

        private List<PAPX> _papxList = new List<PAPX>();
        private List<PAPX> _overFlow;


        public PAPFormattedDiskPage(byte[] dataStream):this()
        {
            
        }

        public PAPFormattedDiskPage()
        {
        }
        /**
         * Creates a PAPFormattedDiskPage from a 512 byte array
         */
        [Obsolete]
        public PAPFormattedDiskPage(byte[] documentStream, byte[] dataStream, int offset, int fcMin, TextPieceTable tpt)
            : this(documentStream, dataStream, offset, tpt )
        {
        }
        /**
 * Creates a PAPFormattedDiskPage from a 512 byte array
 */
        public PAPFormattedDiskPage(byte[] documentStream, byte[] dataStream,
                int offset, CharIndexTranslator translator)
            :base(documentStream, offset)
        {
            for (int x = 0; x < _crun; x++)
            {
                int bytesStartAt = GetStart(x);
                int bytesEndAt = GetEnd(x);

                int charStartAt = translator.GetCharIndex(bytesStartAt);
                int charEndAt = translator.GetCharIndex(bytesEndAt, charStartAt);

                PAPX papx = new PAPX(charStartAt, charEndAt, GetGrpprl(x), GetParagraphHeight(x), dataStream);
                _papxList.Add(papx);
            }
            _fkp = null;
        }

        /**
         * Fills the queue for writing.
         *
         * @param Filler a List of PAPXs
         */
        public void Fill(List<PAPX> Filler)
        {
            _papxList.AddRange(Filler);
        }

        /**
         * Used when writing out a Word docunment. This method is part of a sequence
         * that is necessary because there is no easy and efficient way to
         * determine the number PAPX's that will fit into one FKP. THe sequence is
         * as follows:
         *
         * Fill()
         * ToArray()
         * GetOverflow()
         *
         * @return The remaining PAPXs that didn't fit into this FKP.
         */
       internal List<PAPX> GetOverflow()
        {
            return _overFlow;
        }

        /**
         * Gets the PAPX at index.
         * @param index The index to get the PAPX for.
         * @return The PAPX at index.
         */
        public PAPX GetPAPX(int index)
        {
            return _papxList[index];
        }

        /**
         * Gets the papx grpprl for the paragraph at index in this fkp.
         *
         * @param index The index of the papx to Get.
         * @return a papx grpprl.
         */
        protected override byte[] GetGrpprl(int index)
        {
            int papxOffset = 2 * LittleEndian.GetUByte(_fkp, _offset + (((_crun + 1) * FC_SIZE) + (index * BX_SIZE)));
            int size = 2 * LittleEndian.GetUByte(_fkp, _offset + papxOffset);
            if (size == 0)
            {
                size = 2 * LittleEndian.GetUByte(_fkp, _offset + ++papxOffset);
            }
            else
            {
                size--;
            }

            byte[] papx = new byte[size];
            Array.Copy(_fkp, _offset + ++papxOffset, papx, 0, size);
            return papx;
        }

        /**
         * Creates a byte array representation of this data structure. Suitable for
         * writing to a Word document.
         *
         * @param fcMin The file offset in the main stream where text begins.
         * @return A byte array representing this data structure.
         */
        internal byte[] ToByteArray(HWPFStream dataStream,
            CharIndexTranslator translator)
        {
            byte[] buf = new byte[512];
            int size = _papxList.Count;
            int grpprlOffset = 0;
            int bxOffset = 0;
            int fcOffset = 0;
            byte[] lastGrpprl = Array.Empty<byte>();

            // total size is currently the size of one FC
            int totalSize = FC_SIZE;

            int index = 0;
            for (; index < size; index++)
            {
                byte[] grpprl = ((PAPX)_papxList[index]).GetGrpprl();
                int grpprlLength = grpprl.Length;

                // is grpprl huge?
                if (grpprlLength > 488)
                {
                    grpprlLength = 8; // set equal to size of sprmPHugePapx grpprl
                }

                // check to see if we have enough room for an FC, a BX, and the grpprl
                // and the 1 byte size of the grpprl.
                int addition = 0;
                if (!Arrays.Equals(grpprl, lastGrpprl))
                {
                    addition = (FC_SIZE + BX_SIZE + grpprlLength + 1);
                }
                else
                {
                    addition = (FC_SIZE + BX_SIZE);
                }

                totalSize += addition;

                // if size is uneven we will have to add one so the first grpprl falls
                // on a word boundary
                if (totalSize > 511 + (index % 2))
                {
                    totalSize -= addition;
                    break;
                }

                // grpprls must fall on word boundaries
                if (grpprlLength % 2 > 0)
                {
                    totalSize += 1;
                }
                else
                {
                    totalSize += 2;
                }
                lastGrpprl = grpprl;
            }

            // see if we couldn't fit some
            if (index != size)
            {
                _overFlow = new List<PAPX>();
                _overFlow.AddRange(_papxList.GetRange(index, size-index));
            }

            // index should equal number of papxs that will be in this fkp now.
            buf[511] = (byte)index;

            bxOffset = (FC_SIZE * index) + FC_SIZE;
            grpprlOffset = 511;

            PAPX papx = null;
            lastGrpprl = Array.Empty<byte>();
            for (int x = 0; x < index; x++)
            {
                papx = _papxList[x];
                byte[] phe = papx.GetParagraphHeight().ToArray();
                byte[] grpprl = papx.GetGrpprl();

                // is grpprl huge?
                if (grpprl.Length > 488)
                {
                    /*
                    // if so do we have storage at GetHugeGrpprloffset()
                    int hugeGrpprlOffset = papx.GetHugeGrpprlOffset();
                    if (hugeGrpprlOffset == -1) // then we have no storage...
                    {
                        throw new InvalidOperationException(
                              "This Paragraph has no dataStream storage.");
                    }
                    // we have some storage...

                    // get the size of the existing storage
                    int maxHugeGrpprlSize = LittleEndian.GetUShort(_dataStream, hugeGrpprlOffset);

                    if (maxHugeGrpprlSize < grpprl.Length - 2)
                    { // grpprl.Length-2 because we don't store the istd
                        throw new InvalidOperationException(
                            "This Paragraph's dataStream storage is too small.");
                    }
                   

                    // store grpprl at hugeGrpprlOffset
                    Array.Copy(grpprl, 2, _dataStream, hugeGrpprlOffset + 2,
                                     grpprl.Length - 2); // grpprl.Length-2 because we don't store the istd
                    LittleEndian.PutUShort(_dataStream, hugeGrpprlOffset, grpprl.Length - 2);
                      */

                    byte[] hugePapx = new byte[grpprl.Length - 2];
                    System.Array.Copy(grpprl, 2, hugePapx, 0, grpprl.Length - 2);
                    int dataStreamOffset = dataStream.Offset;
                    dataStream.Write(hugePapx);

                    // grpprl = grpprl Containing only a sprmPHugePapx2
                    int istd = LittleEndian.GetUShort(grpprl, 0);
                    grpprl = new byte[8];
                    LittleEndian.PutUShort(grpprl, 0, istd);
                    LittleEndian.PutUShort(grpprl, 2, 0x6646); // sprmPHugePapx2
                    LittleEndian.PutInt(grpprl, 4, dataStreamOffset);
                }

                bool same = Arrays.Equals(lastGrpprl, grpprl);
                if (!same)
                {
                    grpprlOffset -= (grpprl.Length + (2 - grpprl.Length % 2));
                    grpprlOffset -= (grpprlOffset % 2);
                }
                LittleEndian.PutInt(buf, fcOffset, translator.GetByteIndex(papx.Start));
                buf[bxOffset] = (byte)(grpprlOffset / 2);
                Array.Copy(phe, 0, buf, bxOffset + 1, phe.Length);

                // refer to the section on PAPX in the spec. Places a size on the front
                // of the PAPX. Has to do with how the grpprl stays on word
                // boundaries.
                if (!same)
                {
                    int copyOffset = grpprlOffset;
                    if ((grpprl.Length % 2) > 0)
                    {
                        buf[copyOffset++] = (byte)((grpprl.Length + 1) / 2);
                    }
                    else
                    {
                        buf[++copyOffset] = (byte)((grpprl.Length) / 2);
                        copyOffset++;
                    }
                    Array.Copy(grpprl, 0, buf, copyOffset, grpprl.Length);
                    lastGrpprl = grpprl;
                }

                bxOffset += BX_SIZE;
                fcOffset += FC_SIZE;

            }

            LittleEndian.PutInt(buf, fcOffset, translator.GetByteIndex(papx.End));
            return buf;
        }

        /**
         * Used to get the ParagraphHeight of a PAPX at a particular index.
         * @param index
         * @return The ParagraphHeight
         */
        private ParagraphHeight GetParagraphHeight(int index)
        {
            int pheOffset = _offset + 1 + (((_crun + 1) * 4) + (index * 13));

            ParagraphHeight phe = new ParagraphHeight(_fkp, pheOffset);

            return phe;
        }
    }
}

