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


namespace NPOI.HWPF.UserModel
{
    using NPOI.HWPF.Model;
    using NPOI.HWPF.SPRM;
    using System;
    /**
     * This class is used to create a list in a Word document. It is used in
     * conjunction with {@link
     * NPOI.HWPF.HWPFDocument#registerList(HWPFList) registerList} in
     * {@link NPOI.HWPF.HWPFDocument HWPFDocument}.
     *
     * In Word, lists are not ranged entities, meaning you can't actually add one
     * to the document. Lists only act as properties for list entries. Once you
     * register a list, you can add list entries to a document that are a part of
     * the list.
     *
     * The only benefit of this that I see, is that you can add a list entry
     * anywhere in the document and continue numbering from the previous list.
     *
     * @author Ryan Ackley
     */
    public class HWPFList
    {
        private ListData _listData;
        private ListFormatOverride _override;
        private bool _registered;
        private StyleSheet _styleSheet;

        /**
         *
         * @param numbered true if the list should be numbered; false if it should be
         *        bulleted.
         * @param styleSheet The document's stylesheet.
         */
        public HWPFList(bool numbered, StyleSheet styleSheet)
        {
            _listData = new ListData((int)(new Random((int)DateTime.Now.Ticks).Next(0,100)/100 * DateTime.Now.Millisecond), numbered);
            _override = new ListFormatOverride(_listData.GetLsid());
            _styleSheet = styleSheet;
        }

        /**
         * Sets the character properties of the list numbers.
         *
         * @param level the level number that the properties should apply to.
         * @param chp The character properties.
         */
        public void SetLevelNumberProperties(int level, CharacterProperties chp)
        {
            ListLevel listLevel = _listData.GetLevel(level);
            int styleIndex = _listData.GetLevelStyle(level);
            CharacterProperties base1 = _styleSheet.GetCharacterStyle(styleIndex);

            byte[] grpprl = CharacterSprmCompressor.CompressCharacterProperty(chp, base1);
            listLevel.SetNumberProperties(grpprl);
        }

        /**
         * Sets the paragraph properties for a particular level of the list.
         *
         * @param level The level number.
         * @param pap The paragraph properties
         */
        public void SetLevelParagraphProperties(int level, ParagraphProperties pap)
        {
            ListLevel listLevel = _listData.GetLevel(level);
            int styleIndex = _listData.GetLevelStyle(level);
            ParagraphProperties base1 = _styleSheet.GetParagraphStyle(styleIndex);

            byte[] grpprl = ParagraphSprmCompressor.CompressParagraphProperty(pap, base1);
            listLevel.SetLevelProperties(grpprl);
        }

        public void SetLevelStyle(int level, int styleIndex)
        {
            _listData.SetLevelStyle(level, styleIndex);
        }

        public ListData GetListData()
        {
            return _listData;
        }

        public ListFormatOverride GetOverride()
        {
            return _override;
        }

    }

}