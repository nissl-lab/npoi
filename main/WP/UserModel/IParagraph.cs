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
     * This class represents a paragraph, made up of one or more
     *  Runs of text.
     */
    public interface IParagraph
    {
        // Tables work very differently between the formats
        //  public bool IsInTable();
        //  public bool IsTableRowEnd();
        //  int GetTableLevel();

        // TODO Implement justifaction in XWPF
        //  int GetJustification();
        //  public void SetJustification(byte jc);

        // TODO Expose the different page break related things,
        //  XWPF currently doesn't have the full Set
        /*
            public bool keepOnPage();
            public void SetKeepOnPage(bool fKeep);

            public bool keepWithNext();
            public void SetKeepWithNext(bool fKeepFollow);

            public bool pageBreakBefore();
            public void SetPageBreakBefore(bool fPageBreak);

            public bool IsSideBySide();
            public void SetSideBySide(bool fSideBySide);
        */

        int IndentFromRight { get; set; }

        int IndentFromLeft { get; set; }

        int FirstLineIndent { get; set; }

        /*
            public bool IsLineNotNumbered();
            public void SetLineNotNumbered(bool fNoLnn);

            public bool IsAutoHyphenated();
            public void SetAutoHyphenated(bool autoHyph);

            public bool IsWidowControlled();
            public void SetWidowControl(bool widowControl);

            int GetSpacingBefore();
            public void SetSpacingBefore(int before);

            int GetSpacingAfter();
            public void SetSpacingAfter(int After);
        */

        //  public LineSpacingDescriptor GetLineSpacing();
        //  public void SetLineSpacing(LineSpacingDescriptor lspd);

        int FontAlignment { get; set; }

        bool IsWordWrapped { get; set; }

        /*
            public bool IsVertical();
            public void SetVertical(bool vertical);

            public bool IsBackward();
            public void SetBackward(bool bward);
        */

        // TODO Make the HWPF and XWPF interface wrappers compatible for these
        /*
            public BorderCode GetTopBorder();
            public void SetTopBorder(BorderCode top);
            public BorderCode GetLeftBorder();
            public void SetLeftBorder(BorderCode left);
            public BorderCode GetBottomBorder();
            public void SetBottomBorder(BorderCode bottom);
            public BorderCode GetRightBorder();
            public void SetRightBorder(BorderCode right);
            public BorderCode GetBarBorder();
            public void SetBarBorder(BorderCode bar);

            public ShadingDescriptor GetShading();
            public void SetShading(ShadingDescriptor shd);
        */

        /**
         * Returns the ilfo, an index to the document's hpllfo, which
         *  describes the automatic number formatting of the paragraph.
         * A value of zero means it isn't numbered.
         */
        //    int GetIlfo();

        /**
         * Returns the multi-level indent for the paragraph. Will be
         *  zero for non-list paragraphs, and the first level of any
         *  list. Subsequent levels in hold values 1-8.
         */
        //    int GetIlvl();

        /**
         * Returns the heading level (1-8), or 9 if the paragraph
         *  isn't in a heading style.
         */
        //    int GetLvl();

        /**
         * Returns number of tabs stops defined for paragraph. Must be >= 0 and <=
         * 64.
         * 
         * @return number of tabs stops defined for paragraph. Must be >= 0 and <=
         *         64
         */
        //    int GetTabStopsNumber();

        /**
         * Returns array of positions of itbdMac tab stops
         * 
         * @return array of positions of itbdMac tab stops
         */
        //    int[] GetTabStopsPositions();
    }

}