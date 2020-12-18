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
namespace NPOI.HWPF.Converter
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using NPOI.HWPF;
    using System.Text;
    using NPOI.HWPF.UserModel;
    public class WordToFoUtils : AbstractWordUtils
    {
        public static void CompactInlines(XmlElement blockXmlElement)
        {
            CompactChildNodesR(blockXmlElement, "fo:inline");
        }

        public static void SetBold(XmlElement XmlElement, bool bold)
        {
            XmlElement.SetAttribute("font-weight", bold ? "bold" : "normal");
        }

        public static void SetBorder(XmlElement XmlElement, BorderCode borderCode, String where)
        {
            if (XmlElement == null)
                throw new ArgumentException("XmlElement is null");

            if (borderCode == null || borderCode.BorderType == 0)
                return;

            if (string.IsNullOrEmpty(where))
            {
                XmlElement.SetAttribute("border-style", GetBorderType(borderCode));
                XmlElement.SetAttribute("border-color",
                        GetColor(borderCode.Color));
                XmlElement.SetAttribute("border-width", GetBorderWidth(borderCode));
            }
            else
            {
                XmlElement.SetAttribute("border-" + where + "-style",
                        GetBorderType(borderCode));
                XmlElement.SetAttribute("border-" + where + "-color",
                        GetColor(borderCode.Color));
                XmlElement.SetAttribute("border-" + where + "-width",
                        GetBorderWidth(borderCode));
            }
        }

        public static void SetCharactersProperties(CharacterRun characterRun, XmlElement inline)
        {
            StringBuilder textDecorations = new StringBuilder();

            SetBorder(inline, characterRun.GetBorder(), string.Empty);

            if (characterRun.GetIco24() != -1)
            {
                inline.SetAttribute("color", GetColor24(characterRun.GetIco24()));
            }
            int opacity = (int)(characterRun.GetIco24() & 0xFF000000L) >> 24;
            if (opacity != 0 && opacity != 0xFF)
            {
                inline.SetAttribute("opacity",
                        GetOpacity(characterRun.GetIco24()));
            }
            if (characterRun.IsCapitalized())
            {
                inline.SetAttribute("text-transform", "uppercase");
            }
            if (characterRun.isHighlighted())
            {
                inline.SetAttribute("background-color",
                        GetColor(characterRun.GetHighlightedColor()));
            }
            if (characterRun.IsStrikeThrough())
            {
                if (textDecorations.Length > 0)
                    textDecorations.Append(" ");
                textDecorations.Append("line-through");
            }
            if (characterRun.IsShadowed())
            {
                inline.SetAttribute("text-shadow", characterRun.GetFontSize() / 24 + "pt");
            }
            if (characterRun.IsSmallCaps())
            {
                inline.SetAttribute("font-variant", "small-caps");
            }
            if (characterRun.GetSubSuperScriptIndex() == 1)
            {
                inline.SetAttribute("baseline-shift", "super");
                inline.SetAttribute("font-size", "smaller");
            }
            if (characterRun.GetSubSuperScriptIndex() == 2)
            {
                inline.SetAttribute("baseline-shift", "sub");
                inline.SetAttribute("font-size", "smaller");
            }
            if (characterRun.GetUnderlineCode() > 0)
            {
                if (textDecorations.Length > 0)
                    textDecorations.Append(" ");
                textDecorations.Append("underline");
            }
            if (characterRun.IsVanished())
            {
                inline.SetAttribute("visibility", "hidden");
            }
            if (textDecorations.Length > 0)
            {
                inline.SetAttribute("text-decoration", textDecorations.ToString());
            }
        }

        public static void SetFontFamily(XmlElement XmlElement, String fontFamily)
        {
            if (string.IsNullOrEmpty(fontFamily))
                return;

            XmlElement.SetAttribute("font-family", fontFamily);
        }

        public static void SetFontSize(XmlElement XmlElement, int fontSize)
        {
            XmlElement.SetAttribute("font-size", fontSize.ToString());
        }

        public static void SetIndent(Paragraph paragraph, XmlElement block)
        {
            if (paragraph.GetFirstLineIndent() != 0)
            {
                block.SetAttribute(
                        "text-indent",
                        (paragraph.GetFirstLineIndent() / TWIPS_PER_PT).ToString() + "pt");
            }
            if (paragraph.GetIndentFromLeft() != 0)
            {
                block.SetAttribute(
                        "start-indent",
                        (paragraph.GetIndentFromLeft() / TWIPS_PER_PT).ToString() + "pt");
            }
            if (paragraph.GetIndentFromRight() != 0)
            {
                block.SetAttribute(
                        "end-indent",
                        (paragraph.GetIndentFromRight() / TWIPS_PER_PT).ToString() + "pt");
            }
            if (paragraph.GetSpacingBefore() != 0)
            {
                block.SetAttribute(
                        "space-before",
                        (paragraph.GetSpacingBefore() / TWIPS_PER_PT).ToString() + "pt");
            }
            if (paragraph.GetSpacingAfter() != 0)
            {
                block.SetAttribute("space-after",
                        (paragraph.GetSpacingAfter() / TWIPS_PER_PT) + "pt");
            }
        }

        public static void SetItalic(XmlElement XmlElement, bool italic)
        {
            XmlElement.SetAttribute("font-style", italic ? "italic" : "normal");
        }

        public static void SetJustification(Paragraph paragraph,
                 XmlElement XmlElement)
        {
            String justification = GetJustification(paragraph.GetJustification());
            if (!string.IsNullOrEmpty(justification))
                XmlElement.SetAttribute("text-align", justification);
        }

        public static void SetLanguage(CharacterRun characterRun, XmlElement inline)
        {
            if (characterRun.getLanguageCode() != 0)
            {
                String language = GetLanguage(characterRun.getLanguageCode());
                if (!string.IsNullOrEmpty(language))
                    inline.SetAttribute("language", language);
            }
        }

        public static void SetParagraphProperties(Paragraph paragraph, XmlElement block)
        {
            SetIndent(paragraph, block);
            SetJustification(paragraph, block);

            SetBorder(block, paragraph.GetBottomBorder(), "bottom");
            SetBorder(block, paragraph.GetLeftBorder(), "left");
            SetBorder(block, paragraph.GetRightBorder(), "right");
            SetBorder(block, paragraph.GetTopBorder(), "top");

            if (paragraph.PageBreakBefore())
            {
                block.SetAttribute("break-before", "page");
            }

            block.SetAttribute("hyphenate", paragraph.IsAutoHyphenated.ToString());

            if (paragraph.KeepOnPage())
            {
                block.SetAttribute("keep-together.within-page", "always");
            }

            if (paragraph.KeepWithNext())
            {
                block.SetAttribute("keep-with-next.within-page", "always");
            }

            block.SetAttribute("linefeed-treatment", "preserve");
            block.SetAttribute("white-space-collapse", "false");
        }

        public static void SetPictureProperties(Picture picture, XmlElement graphicXmlElement)
        {
            int horizontalScale = picture.HorizontalScalingFactor;
            int verticalScale = picture.VerticalScalingFactor;

            if (horizontalScale > 0)
            {
                graphicXmlElement
                        .SetAttribute("content-width", ((picture.DxaGoal
                                * horizontalScale / 1000) / TWIPS_PER_PT)
                                + "pt");
            }
            else
                graphicXmlElement.SetAttribute("content-width",
                        (picture.DxaGoal / TWIPS_PER_PT) + "pt");

            if (verticalScale > 0)
                graphicXmlElement
                        .SetAttribute("content-height", ((picture.DyaGoal
                                * verticalScale / 1000) / TWIPS_PER_PT)
                                + "pt");
            else
                graphicXmlElement.SetAttribute("content-height",
                        (picture.DyaGoal / TWIPS_PER_PT) + "pt");

            if (horizontalScale <= 0 || verticalScale <= 0)
            {
                graphicXmlElement.SetAttribute("scaling", "uniform");
            }
            else
            {
                graphicXmlElement.SetAttribute("scaling", "non-uniform");
            }

            graphicXmlElement.SetAttribute("vertical-align", "text-bottom");

            if (picture.DyaCropTop != 0 || picture.DxaCropRight != 0
                    || picture.DyaCropBottom != 0
                    || picture.DxaCropLeft != 0)
            {
                int rectTop = picture.DyaCropTop / TWIPS_PER_PT;
                int rectRight = picture.DxaCropRight / TWIPS_PER_PT;
                int rectBottom = picture.DyaCropBottom / TWIPS_PER_PT;
                int rectLeft = picture.DxaCropLeft / TWIPS_PER_PT;
                graphicXmlElement.SetAttribute("clip", "rect(" + rectTop + "pt, "
                        + rectRight + "pt, " + rectBottom + "pt, " + rectLeft
                        + "pt)");
                graphicXmlElement.SetAttribute("overflow", "hidden");
            }
        }

        public static void SetTableCellProperties(TableRow tableRow,
                TableCell tableCell, XmlElement XmlElement, bool toppest,
                bool bottomest, bool leftest, bool rightest)
        {
            XmlElement.SetAttribute("width", (tableCell.GetWidth() / TWIPS_PER_INCH)
                    + "in");
            XmlElement.SetAttribute("padding-start",
                    (tableRow.GetGapHalf() / TWIPS_PER_INCH) + "in");
            XmlElement.SetAttribute("padding-end",
                    (tableRow.GetGapHalf() / TWIPS_PER_INCH) + "in");

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

            SetBorder(XmlElement, bottom, "bottom");
            SetBorder(XmlElement, left, "left");
            SetBorder(XmlElement, right, "right");
            SetBorder(XmlElement, top, "top");
        }

        public static void SetTableRowProperties(TableRow tableRow, XmlElement tableRowXmlElement)
        {
            if (tableRow.GetRowHeight() > 0)
            {
                tableRowXmlElement.SetAttribute("height",
                        (tableRow.GetRowHeight() / TWIPS_PER_INCH) + "in");
            }
            if (!tableRow.cantSplit())
            {
                tableRowXmlElement.SetAttribute("keep-together.within-column",
                        "always");
            }
        }

    }
}