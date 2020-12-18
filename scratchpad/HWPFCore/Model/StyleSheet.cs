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
    using System;
    using NPOI.Util;
    using NPOI.HWPF.UserModel;
    using NPOI.HWPF.SPRM;
    using NPOI.HWPF.Model.IO;

    /**
     * Represents a document's stylesheet. A word documents formatting is stored as
     * compressed styles that are based on styles Contained in the stylesheet. This
     * class also Contains static utility functions to uncompress different
     * formatting properties.
     *
     * @author Ryan Ackley
     */
    public class StyleSheet
    {

        public const int NIL_STYLE = 4095;
        private const int PAP_TYPE = 1;
        private const int CHP_TYPE = 2;
        private const int SEP_TYPE = 4;
        private const int TAP_TYPE = 5;


        private static ParagraphProperties NIL_PAP = new ParagraphProperties();
        private static CharacterProperties NIL_CHP = new CharacterProperties();

        private int _stshiLength;
        private int _baseLength;
        private int _flags;
        private int _maxIndex;
        private int _maxFixedIndex;
        private int _stylenameVersion;
        private int[] _rgftc;

        StyleDescription[] _styleDescriptions;

        /**
         * StyleSheet constructor. Loads a document's stylesheet information,
         *
         * @param tableStream A byte array Containing a document's raw stylesheet
         *        info. Found by using FileInformationBlock.GetFcStshf() and
         *        FileInformationBLock.GetLcbStshf()
         */
        public StyleSheet(byte[] tableStream, int offset)
        {
            int startoffset = offset;
            _stshiLength = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            int stdCount = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _baseLength = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _flags = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _maxIndex = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _maxFixedIndex = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _stylenameVersion = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;

            _rgftc = new int[3];
            _rgftc[0] = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _rgftc[1] = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;
            _rgftc[2] = LittleEndian.GetShort(tableStream, offset);
            offset += LittleEndianConsts.SHORT_SIZE;

            offset = startoffset + LittleEndianConsts.SHORT_SIZE + _stshiLength;
            _styleDescriptions = new StyleDescription[stdCount];
            for (int x = 0; x < stdCount; x++)
            {
                int stdSize = LittleEndian.GetShort(tableStream, offset);
                //get past the size
                offset += 2;
                if (stdSize > 0)
                {
                    //byte[] std = new byte[stdSize];

                    StyleDescription aStyle = new StyleDescription(tableStream,
                      _baseLength, offset, true);

                    _styleDescriptions[x] = aStyle;
                }

                offset += stdSize;

            }
            for (int x = 0; x < _styleDescriptions.Length; x++)
            {
                if (_styleDescriptions[x] != null)
                {
                    CreatePap(x);
                    CreateChp(x);
                }
            }
        }

        public void WriteTo(HWPFStream out1)
        {
            int offset = 0;
            // add two bytes so we can prepend the stylesheet w/ its size
            byte[] buf = new byte[_stshiLength + 2];
            LittleEndian.PutShort(buf, offset, (short)_stshiLength);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, (short)_styleDescriptions.Length);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, (short)_baseLength);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, (short)_flags);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, (short)_maxIndex);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, (short)_maxFixedIndex);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, (short)_stylenameVersion);
            offset += LittleEndianConsts.SHORT_SIZE;

            LittleEndian.PutShort(buf, offset, (short)_rgftc[0]);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, (short)_rgftc[1]);
            offset += LittleEndianConsts.SHORT_SIZE;
            LittleEndian.PutShort(buf, offset, (short)_rgftc[2]);

            out1.Write(buf);

            byte[] sizeHolder = new byte[2];
            for (int x = 0; x < _styleDescriptions.Length; x++)
            {
                if (_styleDescriptions[x] != null)
                {
                    byte[] std = _styleDescriptions[x].ToArray();

                    // adjust the size so it is always on a word boundary
                    LittleEndian.PutShort(sizeHolder, (short)((std.Length) + (std.Length % 2)));
                    out1.Write(sizeHolder);
                    out1.Write(std);

                    // Must always start on a word boundary.
                    if (std.Length % 2 == 1)
                    {
                        out1.Write('\0');
                    }
                }
                else
                {
                    sizeHolder[0] = 0;
                    sizeHolder[1] = 0;
                    out1.Write(sizeHolder);
                }
            }
        }

   
        public override bool Equals(Object o)
        {
            StyleSheet ss = (StyleSheet)o;

            if (ss._baseLength == _baseLength && ss._flags == _flags &&
                ss._maxFixedIndex == _maxFixedIndex && ss._maxIndex == _maxIndex &&
                ss._rgftc[0] == _rgftc[0] && ss._rgftc[1] == _rgftc[1] &&
                ss._rgftc[2] == _rgftc[2] && ss._stshiLength == _stshiLength &&
                ss._stylenameVersion == _stylenameVersion)
            {
                if (ss._styleDescriptions.Length == _styleDescriptions.Length)
                {
                    for (int x = 0; x < _styleDescriptions.Length; x++)
                    {
                        // check for null
                        if (ss._styleDescriptions[x] != _styleDescriptions[x])
                        {
                            // check for Equality
                            if (!ss._styleDescriptions[x].Equals(_styleDescriptions[x]))
                            {
                                return false;
                            }
                        }
                    }
                    return true;
                }
            }
            return false;
        }
        /**
         * Creates a PartagraphProperties object from a papx stored in the
         * StyleDescription at the index istd in the StyleDescription array. The PAP
         * is placed in the StyleDescription at istd after its been Created. Not
         * every StyleDescription will contain a papx. In these cases this function
         * does nothing
         *
         * @param istd The index of the StyleDescription to create the
         *        ParagraphProperties  from (and also place the finished PAP in)
         */
        private void CreatePap(int istd)
        {
            StyleDescription sd = _styleDescriptions[istd];
            ParagraphProperties pap = sd.GetPAP();
            byte[] papx = sd.GetPAPX();
            int baseIndex = sd.GetBaseStyle();
            if (pap == null && papx != null)
            {
                ParagraphProperties parentPAP = new ParagraphProperties();
                if (baseIndex != NIL_STYLE)
                {

                    parentPAP = _styleDescriptions[baseIndex].GetPAP();
                    if (parentPAP == null)
                    {
                        if (baseIndex == istd)
                        {
                            // Oh dear, style claims that it is its own parent
                            throw new InvalidOperationException("Pap style " + istd + " claimed to have itself as its parent, which isn't allowed");
                        }
                        // Create the parent style
                        CreatePap(baseIndex);
                        parentPAP = _styleDescriptions[baseIndex].GetPAP();
                    }

                }

                pap = ParagraphSprmUncompressor.UncompressPAP(parentPAP, papx, 2);
                sd.SetPAP(pap);
            }
        }
        /**
         * Creates a CharacterProperties object from a chpx stored in the
         * StyleDescription at the index istd in the StyleDescription array. The
         * CharacterProperties object is placed in the StyleDescription at istd after
         * its been Created. Not every StyleDescription will contain a chpx. In these
         * cases this function does nothing.
         *
         * @param istd The index of the StyleDescription to create the
         *        CharacterProperties object from.
         */
        private void CreateChp(int istd)
        {
            StyleDescription sd = _styleDescriptions[istd];
            CharacterProperties chp = sd.GetCHP();
            byte[] chpx = sd.GetCHPX();
            int baseIndex = sd.GetBaseStyle();

            if (baseIndex == istd)
            {
                // Oh dear, this isn't allowed...
                // The word file seems to be corrupted
                // Switch to using the nil style so that
                //  there's a chance we can read it
                baseIndex = NIL_STYLE;
            }

            // Build and decompress the Chp if required 
            if (chp == null && chpx != null)
            {
                CharacterProperties parentCHP = new CharacterProperties();
                if (baseIndex != NIL_STYLE)
                {

                    parentCHP = _styleDescriptions[baseIndex].GetCHP();
                    if (parentCHP == null)
                    {
                        CreateChp(baseIndex);
                        parentCHP = _styleDescriptions[baseIndex].GetCHP();
                    }

                }

                chp = CharacterSprmUncompressor.UncompressCHP(parentCHP, chpx, 0);
                sd.SetCHP(chp);
            }
        }

        /**
         * Gets the number of styles in the style sheet.
         * @return The number of styles in the style sheet.
         */
        public int NumStyles()
        {
            return _styleDescriptions.Length;
        }

        /**
         * Gets the StyleDescription at index x.
         *
         * @param x the index of the desired StyleDescription.
         */
        public StyleDescription GetStyleDescription(int x)
        {
            return _styleDescriptions[x];
        }

        public CharacterProperties GetCharacterStyle(int x)
        {
            if (x == NIL_STYLE)
            {
                return NIL_CHP;
            }
            if (x >= _styleDescriptions.Length)
            {
                return NIL_CHP;
            }
            return (_styleDescriptions[x] != null ? _styleDescriptions[x].GetCHP() : NIL_CHP);
        }

        public ParagraphProperties GetParagraphStyle(int x)
        {
            if (x == NIL_STYLE)
            {
                return NIL_PAP;
            }

            if (x >= _styleDescriptions.Length)
            {
                return NIL_PAP;
            }

            if (_styleDescriptions[x] == null)
            {
                return NIL_PAP;
            }

            if (_styleDescriptions[x].GetPAP() == null)
            {
                return NIL_PAP;
            }

            return _styleDescriptions[x].GetPAP();
        }

    }


}