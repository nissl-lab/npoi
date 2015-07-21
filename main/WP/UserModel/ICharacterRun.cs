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

namespace NPOI.WP.UserModel
{
    using System;

    /**
     * This class represents a run of text that share common properties.
     */
    public interface ICharacterRun
    {
        bool IsBold { get; set; }

        bool IsItalic { get; set; }
        bool IsSmallCaps { get; set; }

        bool IsCapitalized { get; set; }

        bool IsStrikeThrough { get; set; }
        bool IsDoubleStrikeThrough { get; set; }

        bool IsShadowed { get; set; }

        bool IsEmbossed { get; set; }

        bool IsImprinted { get; set; }

        int FontSize { get; set; }

        int CharacterSpacing { get; set; }

        int Kerning { get; set; }

        String FontName { get; }

        /**
         * @return The text of the Run, including any tabs/spaces/etc
         */
        String Text { get; }

        // HWPF uses indexes, XWPF special
        //    int GetUnderlineCode();
        //    public void SetUnderlineCode(int kul);

        // HWPF uses indexes, XWPF special vertical alignments
        //    public short GetSubSuperScriptIndex();
        //    public void SetSubSuperScriptIndex(short iss);

        // HWPF uses indexes, XWPF special vertical alignments
        //    int GetVerticalOffset();
        //    public void SetVerticalOffset(int hpsPos);

        // HWPF has colour indexes, XWPF colour names
        //    int GetColor();
        //    public void SetColor(int color);

        // TODO Review these, and add to XWPFRun if possible
        /*
            bool IsFldVanished();
            public void SetFldVanish(bool fldVanish);
    
            bool IsOutlined();
            public void SetOutline(bool outlined);
    
            bool IsVanished();
            public void SetVanished(bool vanish);

            bool IsMarkedDeleted();
            public void markDeleted(bool mark);

            bool IsMarkedInserted();
            public void markInserted(bool mark);
        */
    }

}