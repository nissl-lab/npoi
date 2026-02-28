/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HWPF.UserModel;
using System.Xml;

namespace NPOI.HWPF.Converter
{
    /// <summary>
    /// Beta
    /// </summary>
    public class WordToHtmlUtils : AbstractWordUtils
    {
        public static void AddBold(bool bold, StringBuilder style)
        {
            style.Append("font-weight:" + (bold ? "bold" : "normal") + ";");
        }

        public static void AddBorder(BorderCode borderCode, String where, StringBuilder style)
        {
            if (borderCode == null || borderCode.IsEmpty)
                return;

            if (string.IsNullOrEmpty(where))
            {
                style.Append("border:");
            }
            else
            {
                style.Append("border-");
                style.Append(where);
            }

            style.Append(":");
            if (borderCode.LineWidth < 8)
                style.Append("thin");
            else
                style.Append(GetBorderWidth(borderCode));
            style.Append(' ');
            style.Append(GetBorderType(borderCode));
            style.Append(' ');
            style.Append(GetColor(borderCode.Color));
            style.Append(';');
        }
        public static void AddCharactersProperties(CharacterRun characterRun, StringBuilder style)
        {
            AddBorder(characterRun.GetBorder(), string.Empty, style);

            if (characterRun.IsCapitalized())
            {
                style.Append("text-transform:uppercase;");
            }
            if (characterRun.GetIco24() != -1)
            {
                style.Append("color:" + GetColor24(characterRun.GetIco24()) + ";");
            }
            if (characterRun.IsHighlighted())
            {
                style.Append("background-color:" + GetColor(characterRun.GetHighlightedColor()) + ";");
            }
            if (characterRun.IsStrikeThrough())
            {
                style.Append("text-decoration:line-through;");
            }
            if (characterRun.IsShadowed())
            {
                style.Append("text-shadow:" + characterRun.GetFontSize() / 24 + "pt;");
            }
            if (characterRun.IsSmallCaps())
            {
                style.Append("font-variant:small-caps;");
            }
            if (characterRun.GetSubSuperScriptIndex() == 1)
            {
                style.Append("vertical-align:super;");
                style.Append("font-size:smaller;");
            }
            if (characterRun.GetSubSuperScriptIndex() == 2)
            {
                style.Append("vertical-align:sub;");
                style.Append("font-size:smaller;");
            }
            if (characterRun.GetUnderlineCode() > 0)
            {
                style.Append("text-decoration:underline;");
            }
            if (characterRun.IsVanished())
            {
                style.Append("visibility:hidden;");
            }
        }

        public static void AddFontFamily(String fontFamily, StringBuilder style)
        {
            if (string.IsNullOrEmpty(fontFamily))
                return;

            style.Append("font-family:" + fontFamily + ";");
        }

        public static void AddFontSize(int fontSize, StringBuilder style)
        {
            style.Append("font-size:" + fontSize + "pt;");
        }

        public static void AddIndent(Paragraph paragraph, StringBuilder style)
        {
            AddIndent(style, "text-indent", paragraph.GetFirstLineIndent());
            AddIndent(style, "start-indent", paragraph.GetIndentFromLeft());
            AddIndent(style, "end-indent", paragraph.GetIndentFromRight());
            AddIndent(style, "space-before", paragraph.GetSpacingBefore());
            AddIndent(style, "space-after", paragraph.GetSpacingAfter());
        }

        private static void AddIndent(StringBuilder style, String cssName, int twipsValue)
        {
            if (twipsValue == 0)
                return;

            style.Append(cssName + ":" + (twipsValue / TWIPS_PER_PT) + "pt;");
        }

        public static void AddJustification(Paragraph paragraph, StringBuilder style)
        {
            String justification = GetJustification(paragraph.GetJustification());
            if (!string.IsNullOrEmpty(justification))
                style.Append("text-align:" + justification + ";");
        }

        public static void AddParagraphProperties(Paragraph paragraph, StringBuilder style)
        {
            AddIndent(paragraph, style);
            AddJustification(paragraph, style);

            AddBorder(paragraph.GetBottomBorder(), "bottom", style);
            AddBorder(paragraph.GetLeftBorder(), "left", style);
            AddBorder(paragraph.GetRightBorder(), "right", style);
            AddBorder(paragraph.GetTopBorder(), "top", style);

            if (paragraph.PageBreakBefore())
            {
                style.Append("break-before:page;");
            }

            style.Append("hyphenate:" + (paragraph.IsAutoHyphenated ? "auto" : "none") + ";");

            if (paragraph.KeepOnPage())
            {
                style.Append("keep-together.within-page:always;");
            }

            if (paragraph.KeepWithNext())
            {
                style.Append("keep-with-next.within-page:always;");
            }
        }

        public static void AddTableCellProperties(TableRow tableRow,
                TableCell tableCell, bool toppest, bool bottomest,
                bool leftest, bool rightest, StringBuilder style)
        {
            style.Append("width:" + (tableCell.GetWidth() / TWIPS_PER_INCH) + "in;");
            style.Append("padding-start:" + (tableRow.GetGapHalf() / TWIPS_PER_INCH) + "in;");
            style.Append("padding-end:" + (tableRow.GetGapHalf() / TWIPS_PER_INCH) + "in;");

            BorderCode top = tableCell.GetBrcTop() != null
                    && tableCell.GetBrcTop().BorderType != 0 ? tableCell
                    .GetBrcTop() : toppest ? tableRow.GetTopBorder() : tableRow
                    .GetHorizontalBorder();
            BorderCode bottom = tableCell.GetBrcBottom() != null
                    && tableCell.GetBrcBottom().BorderType != 0 ? tableCell
                    .GetBrcBottom() : bottomest ? tableRow.GetBottomBorder()
                    : tableRow.GetHorizontalBorder();

            BorderCode left = tableCell.GetBrcLeft() != null
                    && tableCell.GetBrcLeft().BorderType != 0 ? tableCell
                    .GetBrcLeft() : leftest ? tableRow.GetLeftBorder() : tableRow
                    .GetVerticalBorder();
            BorderCode right = tableCell.GetBrcRight() != null
                    && tableCell.GetBrcRight().BorderType != 0 ? tableCell
                    .GetBrcRight() : rightest ? tableRow.GetRightBorder()
                    : tableRow.GetVerticalBorder();

            AddBorder(bottom, "bottom", style);
            AddBorder(left, "left", style);
            AddBorder(right, "right", style);
            AddBorder(top, "top", style);
        }

        public static void AddTableRowProperties(TableRow tableRow, StringBuilder style)
        {
            if (tableRow.GetRowHeight() > 0)
            {
                style.Append("height:" + (tableRow.GetRowHeight() / TWIPS_PER_INCH) + "in;");
            }
            if (!tableRow.cantSplit())
            {
                style.Append("keep-together:always;");
            }
        }

        public static void CompactSpans(XmlElement pElement)
        {
            CompactChildNodesR(pElement, "span");
        }
    }
}
