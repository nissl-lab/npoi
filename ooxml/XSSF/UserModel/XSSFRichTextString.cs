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

using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Text;
using System.Collections.Generic;
using NPOI.XSSF.Model;
namespace NPOI.XSSF.UserModel
{

    /**
     * Rich text unicode string.  These strings can have fonts applied to arbitary parts of the string.
     *
     * <p>
     * Most strings in a workbook have formatting applied at the cell level, that is, the entire string in the cell has the
     * same formatting applied. In these cases, the formatting for the cell is stored in the styles part,
     * and the string for the cell can be shared across the workbook. The following code illustrates the example.
     * </p>
     *
     * <blockquote>
     * <pre>
     *     cell1.SetCellValue(new XSSFRichTextString("Apache POI"));
     *     cell2.SetCellValue(new XSSFRichTextString("Apache POI"));
     *     cell3.SetCellValue(new XSSFRichTextString("Apache POI"));
     * </pre>
     * </blockquote>
     * In the above example all three cells will use the same string cached on workbook level.
     *
     * <p>
     * Some strings in the workbook may have formatting applied at a level that is more granular than the cell level.
     * For instance, specific characters within the string may be bolded, have coloring, italicizing, etc.
     * In these cases, the formatting is stored along with the text in the string table, and is treated as
     * a unique entry in the workbook. The following xml and code snippet illustrate this.
     * </p>
     *
     * <blockquote>
     * <pre>
     *     XSSFRichTextString s1 = new XSSFRichTextString("Apache POI");
     *     s1.ApplyFont(boldArial);
     *     cell1.SetCellValue(s1);
     *
     *     XSSFRichTextString s2 = new XSSFRichTextString("Apache POI");
     *     s2.ApplyFont(italicCourier);
     *     cell2.SetCellValue(s2);
     * </pre>
     * </blockquote>
     *
     *
     * @author Yegor Kozlov
     */
    public class XSSFRichTextString : IRichTextString
    {
        private static Regex utfPtrn = new Regex("_x([0-9A-F]{4})_");

        private CT_Rst st;
        private StylesTable styles;

        /**
         * Create a rich text string
         */
        public XSSFRichTextString(String str)
        {
            st = new CT_Rst();
            st.t = str;
            PreserveSpaces(st.t);
        }



        public void SetStylesTableReference(StylesTable stylestable)
        {
            this.styles = stylestable;
            if (st.sizeOfRArray() > 0)
            {
                foreach (CT_RElt r in st.r)
                {
                    CT_RPrElt pr = r.rPr;
                    if (pr != null && pr.SizeOfRFontArray() > 0)
                    {
                        String fontName = pr.GetRFontArray(0).val;
                        if (fontName.StartsWith("#"))
                        {
                            int idx = int.Parse(fontName.Substring(1));
                            XSSFFont font = styles.GetFontAt(idx);
                            pr.rFont = null;
                            SetRunAttributes(font.GetCTFont(), pr);
                        }
                    }
                }
            }
        }
        /**
         * Create empty rich text string and Initialize it with empty string
         */
        public XSSFRichTextString()
        {
            st = new CT_Rst();
        }

        /**
         * Create a rich text string from the supplied XML bean
         */
        public XSSFRichTextString(CT_Rst st)
        {
            this.st = st;
        }

        /**
         * Applies a font to the specified characters of a string.
         *
         * @param startIndex    The start index to apply the font to (inclusive)
         * @param endIndex      The end index to apply the font to (exclusive)
         * @param fontIndex     The font to use.
         */
        public void ApplyFont(int startIndex, int endIndex, short fontIndex)
        {
            XSSFFont font;
            if (styles == null)
            {
                //style table is not Set, remember fontIndex and Set the run properties later,
                //when SetStylesTableReference is called
                font = new XSSFFont();
                font.FontName = ("#" + fontIndex);
            }
            else
            {
                font = styles.GetFontAt(fontIndex);
            }
            ApplyFont(startIndex, endIndex, font);
        }
        internal void ApplyFont(SortedDictionary<int, CT_RPrElt> formats, int startIndex, int endIndex, CT_RPrElt fmt) 
        {
            
            // delete format runs that fit between startIndex and endIndex
            // runs intersecting startIndex and endIndex remain
            //int runStartIdx = 0;
            List<int> toRemoveKeys=new List<int>();
            for (SortedDictionary<int, CT_RPrElt>.KeyCollection.Enumerator it = formats.Keys.GetEnumerator(); it.MoveNext(); )
            {
                int runIdx = it.Current;
                if (runIdx >= startIndex && runIdx < endIndex)
                {
                    toRemoveKeys.Add(runIdx);
                }
            }
            foreach (int key in toRemoveKeys)
            {
                formats.Remove(key);
            }
            if (startIndex > 0 && !formats.ContainsKey(startIndex))
            {
                // If there's a format that starts later in the string, make it start now
                foreach (KeyValuePair<int, CT_RPrElt> entry in formats)
                {
                    if (entry.Key > startIndex)
                    {
                        formats[startIndex] = entry.Value;
                        break;
                    }
                }
            }
            formats[endIndex] = fmt;

            // assure that the range [startIndex, endIndex] consists if a single run
            // there can be two or three runs depending whether startIndex or endIndex
            // intersected existing format runs
            //SortedMap<int, CT_RPrElt> sub = formats.subMap(startIndex, endIndex);
            //while(sub.size() > 1) sub.remove(sub.lastKey());       
        }
        /**
         * Applies a font to the specified characters of a string.
         *
         * @param startIndex    The start index to apply the font to (inclusive)
         * @param endIndex      The end index to apply to font to (exclusive)
         * @param font          The index of the font to use.
         */
        public void ApplyFont(int startIndex, int endIndex, IFont font)
        {
            if (startIndex > endIndex)
                throw new ArgumentException("Start index must be less than end index, but had " + startIndex + " and " + endIndex);
            if (startIndex < 0 || endIndex > Length)
                throw new ArgumentException("Start and end index not in range, but had " + startIndex + " and " + endIndex);
            if (startIndex == endIndex)
                return;

            if (st.sizeOfRArray() == 0 && st.IsSetT())
            {
                //convert <t>string</t> into a text Run: <r><t>string</t></r>
                st.AddNewR().t = (st.t);
                st.unsetT();
            }

            String text = this.String;
            XSSFFont xssfFont = (XSSFFont)font;

            SortedDictionary<int, CT_RPrElt> formats = GetFormatMap(st);
            CT_RPrElt fmt = new CT_RPrElt();
            SetRunAttributes(xssfFont.GetCTFont(), fmt);
            ApplyFont(formats, startIndex, endIndex, fmt);

            CT_Rst newSt = buildCTRst(text, formats);
            st.Set(newSt);



        }
        internal SortedDictionary<int, CT_RPrElt> GetFormatMap(CT_Rst entry)
        {
            int length = 0;
            SortedDictionary<int, CT_RPrElt> formats = new SortedDictionary<int, CT_RPrElt>();
            foreach (CT_RElt r in entry.r)
            {
                String txt = r.t;
                CT_RPrElt fmt = r.rPr;

                length += txt.Length;
                formats[length] = fmt;
            }
            return formats;
        }
        /**
         * Sets the font of the entire string.
         * @param font          The font to use.
         */
        public void ApplyFont(IFont font)
        {
            String text = this.String;
            ApplyFont(0, text.Length, font);
        }

        /**
         * Applies the specified font to the entire string.
         *
         * @param fontIndex  the font to Apply.
         */
        public void ApplyFont(short fontIndex)
        {
            XSSFFont font;
            if (styles == null)
            {
                font = new XSSFFont();
                font.FontName = ("#" + fontIndex);
            }
            else
            {
                font = styles.GetFontAt(fontIndex);
            }
            String text = this.String;
            ApplyFont(0, text.Length, font);
        }

        /**
         * Append new text to this text run and apply the specify font to it
         *
         * @param text  the text to append
         * @param font  the font to apply to the Appended text or <code>null</code> if no formatting is required
         */
        public void Append(String text, XSSFFont font)
        {
            if (st.sizeOfRArray() == 0 && st.IsSetT())
            {
                //convert <t>string</t> into a text Run: <r><t>string</t></r>
                CT_RElt lt = st.AddNewR();
                lt.t = st.t;
                PreserveSpaces(lt.t);
                st.unsetT();
            }
            CT_RElt lt2 = st.AddNewR();
            lt2.t= (text);
            PreserveSpaces(lt2.t);
            CT_RPrElt pr = lt2.AddNewRPr();
            if (font != null) SetRunAttributes(font.GetCTFont(), pr);
        }

        /**
         * Append new text to this text run
         *
         * @param text  the text to append
         */
        public void Append(String text)
        {
            Append(text, null);
        }

        /**
         * Copy font attributes from CTFont bean into CTRPrElt bean
         */
        private void SetRunAttributes(CT_Font ctFont, CT_RPrElt pr)
        {
            if (ctFont.SizeOfBArray() > 0) pr.AddNewB().val = (ctFont.GetBArray(0).val);
            if (ctFont.sizeOfUArray() > 0) pr.AddNewU().val =(ctFont.GetUArray(0).val);
            if (ctFont.sizeOfIArray() > 0) pr.AddNewI().val =(ctFont.GetIArray(0).val);
            if (ctFont.sizeOfColorArray() > 0)
            {
                CT_Color c1 = ctFont.GetColorArray(0);
                CT_Color c2 = pr.AddNewColor();
                if (c1.IsSetAuto())
                {
                    c2.auto = (c1.auto);
                    c2.autoSpecified = true;
                }
                if (c1.IsSetIndexed())
                {
                    c2.indexed = (c1.indexed);
                    c2.indexedSpecified = true;
                }
                if (c1.IsSetRgb())
                {
                    c2.SetRgb(c1.rgb);
                    c2.rgbSpecified = true;
                }
                if (c1.IsSetTheme())
                {
                    c2.theme = (c1.theme);
                    c2.themeSpecified = true;
                }
                if (c1.IsSetTint())
                {
                    c2.tint = (c1.tint);
                    c2.tintSpecified = true;
                }
            }

            if (ctFont.sizeOfSzArray() > 0) pr.AddNewSz().val = (ctFont.GetSzArray(0).val);
            if (ctFont.sizeOfNameArray() > 0) pr.AddNewRFont().val = (ctFont.name.val);
            if (ctFont.sizeOfFamilyArray() > 0) pr.AddNewFamily().val =(ctFont.GetFamilyArray(0).val);
            if (ctFont.sizeOfSchemeArray() > 0) pr.AddNewScheme().val = (ctFont.GetSchemeArray(0).val);
            if (ctFont.sizeOfCharsetArray() > 0) pr.AddNewCharset().val = (ctFont.GetCharsetArray(0).val);
            if (ctFont.sizeOfCondenseArray() > 0) pr.AddNewCondense().val = (ctFont.GetCondenseArray(0).val);
            if (ctFont.sizeOfExtendArray() > 0) pr.AddNewExtend().val = (ctFont.GetExtendArray(0).val);
            if (ctFont.sizeOfVertAlignArray() > 0) pr.AddNewVertAlign().val = (ctFont.GetVertAlignArray(0).val);
            if (ctFont.sizeOfOutlineArray() > 0) pr.AddNewOutline().val =(ctFont.GetOutlineArray(0).val);
            if (ctFont.sizeOfShadowArray() > 0) pr.AddNewShadow().val =(ctFont.GetShadowArray(0).val);
            if (ctFont.sizeOfStrikeArray() > 0) pr.AddNewStrike().val = (ctFont.GetStrikeArray(0).val);
        }

        /**
         * Removes any formatting that may have been applied to the string.
         */
        public void ClearFormatting()
        {
            String text = this.String;
            st.r = (null);
            st.t= (text);
        }

        /**
         * The index within the string to which the specified formatting run applies.
         *
         * @param index     the index of the formatting run
         * @return  the index within the string.
         */
        public int GetIndexOfFormattingRun(int index)
        {
            if (st.sizeOfRArray() == 0) return 0;

            int pos = 0;
            for (int i = 0; i < st.sizeOfRArray(); i++)
            {
                CT_RElt r = st.GetRArray(i);
                if (i == index) return pos;

                pos += r.t.Length;
            }
            return -1;
        }

        /**
         * Returns the number of characters this format run covers.
         *
         * @param index     the index of the formatting run
         * @return  the number of characters this format run covers
         */
        public int GetLengthOfFormattingRun(int index)
        {
            if (st.sizeOfRArray() == 0 || index >= st.sizeOfRArray())
            {
                return -1;
            }

            CT_RElt r = st.GetRArray(index);
            return r.t.Length;
        }

        public String String
        {
            get
            {
                if (st.sizeOfRArray() == 0)
                {
                    return UtfDecode(st.t);
                }
                StringBuilder buf = new StringBuilder();
                foreach (CT_RElt r in st.r)
                {
                    buf.Append(r.t);
                }
                return UtfDecode(buf.ToString());
            }

            set 
            {
                ClearFormatting();
                st.t = value;
                PreserveSpaces(st.t);
            }
        }

        /**
         * Returns the plain string representation.
         */
        public override String ToString()
        {
            return this.String;
        }

        /**
         * Returns the number of characters in this string.
         */
        public int Length
        {
            get
            {
                return this.String.Length;
            }
        }

        /**
         * @return  The number of formatting Runs used.
         */
        public int NumFormattingRuns
        {
            get
            {
                return st.sizeOfRArray();
            }
        }

        /**
         * Gets a copy of the font used in a particular formatting Run.
         *
         * @param index     the index of the formatting run
         * @return  A copy of the  font used or null if no formatting is applied to the specified text Run.
         */
        public IFont GetFontOfFormattingRun(int index)
        {
            if (st.sizeOfRArray() == 0 || index >= st.sizeOfRArray()) return null;

            CT_RElt r = st.GetRArray(index);
            if (r.rPr != null)
            {
                XSSFFont fnt = new XSSFFont(ToCTFont(r.rPr));
                fnt.SetThemesTable(GetThemesTable());
                return fnt;
            }
            return null;
        }

        /**
         * Return a copy of the font in use at a particular index.
         *
         * @param index         The index.
         * @return              A copy of the  font that's currently being applied at that
         *                      index or null if no font is being applied or the
         *                      index is out of range.
         */
        public XSSFFont GetFontAtIndex(int index)
        {
            if (st.sizeOfRArray() == 0) return null;

            int pos = 0;
            for (int i = 0; i < st.sizeOfRArray(); i++)
            {
                CT_RElt r = st.GetRArray(i);
                if (index >= pos && index < pos + r.t.Length)
                {
                    XSSFFont fnt = new XSSFFont(ToCTFont(r.rPr));
                    fnt.SetThemesTable(GetThemesTable());
                    return fnt;
                }

                pos += r.t.Length;
            }
            return null;

        }

        /**
         * Return the underlying xml bean
         */

        public CT_Rst GetCTRst()
        {
            return st;
        }


        /**
         *
         * CTRPrElt --> CTFont adapter
         */
        protected static CT_Font ToCTFont(CT_RPrElt pr)
        {
            CT_Font ctFont = new CT_Font();

            if (pr.SizeOfBArray() > 0) ctFont.AddNewB().val = (pr.GetBArray(0).val);
            if (pr.SizeOfUArray() > 0) ctFont.AddNewU().val = (pr.GetUArray(0).val);
            if (pr.SizeOfIArray() > 0) ctFont.AddNewI().val = (pr.GetIArray(0).val);
            if (pr.SizeOfColorArray() > 0)
            {
                CT_Color c1 = pr.GetColorArray(0);
                CT_Color c2 = ctFont.AddNewColor();
                if (c1.IsSetAuto())
                {
                    c2.auto = (c1.auto);
                    c2.autoSpecified = true;
                }
                if (c1.IsSetIndexed())
                {
                    c2.indexed = (c1.indexed);
                    c2.indexedSpecified = true;
                }
                if (c1.IsSetRgb())
                {
                    c2.SetRgb(c1.GetRgb());
                    c2.rgbSpecified = true;
                }
                if (c1.IsSetTheme())
                {
                    c2.theme = (c1.theme);
                    c2.themeSpecified = true;
                }
                if (c1.IsSetTint())
                {
                    c2.tint = (c1.tint);
                    c2.tintSpecified = true;
                }
            }
 
            if (pr.SizeOfSzArray() > 0) ctFont.AddNewSz().val = (pr.GetSzArray(0).val);
            if (pr.SizeOfRFontArray() > 0) ctFont.AddNewName().val = (pr.GetRFontArray(0).val);
            if (pr.SizeOfFamilyArray() > 0) ctFont.AddNewFamily().val = (pr.GetFamilyArray(0).val);
            if (pr.sizeOfSchemeArray() > 0) ctFont.AddNewScheme().val = (pr.GetSchemeArray(0).val);
            if (pr.sizeOfCharsetArray() > 0) ctFont.AddNewCharset().val = (pr.GetCharsetArray(0).val);
            if (pr.sizeOfCondenseArray() > 0) ctFont.AddNewCondense().val = (pr.GetCondenseArray(0).val);
            if (pr.sizeOfExtendArray() > 0) ctFont.AddNewExtend().val = (pr.GetExtendArray(0).val);
            if (pr.sizeOfVertAlignArray() > 0) ctFont.AddNewVertAlign().val = (pr.GetVertAlignArray(0).val);
            if (pr.sizeOfOutlineArray() > 0) ctFont.AddNewOutline().val = (pr.GetOutlineArray(0).val);
            if (pr.sizeOfShadowArray() > 0) ctFont.AddNewShadow().val = (pr.GetShadowArray(0).val);
            if (pr.sizeOfStrikeArray() > 0) ctFont.AddNewStrike().val = (pr.GetStrikeArray(0).val);

            return ctFont;
        }

        ///**
        // * Add the xml:spaces="preserve" attribute if the string has leading or trailing spaces
        // *
        // * @param xs    the string to check
        // */
        protected static void PreserveSpaces(string xs)
        {
            String text = xs;
            if (text != null && text.Length > 0)
            {
                char firstChar = text[0];
                char lastChar = text[text.Length - 1];
                if (Char.IsWhiteSpace(firstChar) || Char.IsWhiteSpace(lastChar))
                {
                    //XmlCursor c = xs.newCursor();
                    //c.ToNextToken();
                    //c.insertAttributeWithValue(new QName("http://www.w3.org/XML/1998/namespace", "space"), "preserve");
                    //c.dispose();
                }
            }
        }

        /**
         * For all characters which cannot be represented in XML as defined by the XML 1.0 specification,
         * the characters are escaped using the Unicode numerical character representation escape character
         * format _xHHHH_, where H represents a hexadecimal character in the character's value.
         * <p>
         * Example: The Unicode character 0D is invalid in an XML 1.0 document,
         * so it shall be escaped as <code>_x000D_</code>.
         * </p>
         * See section 3.18.9 in the OOXML spec.
         *
         * @param   value the string to decode
         * @return  the decoded string
         */
        static String UtfDecode(String value)
        {
            if (value == null) return null;

            StringBuilder buf = new StringBuilder();
            MatchCollection mc = utfPtrn.Matches(value);
            int idx = 0;
            for (int i = 0; i < mc.Count;i++ )
            {
                    int pos = mc[i].Index;
                    if (pos > idx)
                    {
                        buf.Append(value.Substring(idx, pos-idx));
                    }

                    String code = mc[i].Groups[1].Value;
                    int icode = Int32.Parse(code, System.Globalization.NumberStyles.AllowHexSpecifier);
                    buf.Append((char)icode);

                    idx = mc[i].Index+mc[i].Length;
                }
            buf.Append(value.Substring(idx));
            return buf.ToString();
        }

        public int GetLastKey(SortedDictionary<int, CT_RPrElt>.KeyCollection keys)
        {
            int i=0;
            foreach (int key in keys)
            {
                if (i == keys.Count - 1)
                    return key;
                i++;
            }
            throw new ArgumentOutOfRangeException("GetLastKey failed");
        }

        CT_Rst buildCTRst(String text, SortedDictionary<int, CT_RPrElt> formats)
        {
            if (text.Length != GetLastKey(formats.Keys))
            {
                throw new ArgumentException("Text length was " + text.Length +
                        " but the last format index was " + GetLastKey(formats.Keys));
            }
            CT_Rst st = new CT_Rst();
            int runStartIdx = 0;
            for (SortedDictionary<int, CT_RPrElt>.KeyCollection.Enumerator it = formats.Keys.GetEnumerator(); it.MoveNext(); )
            {
                int runEndIdx = it.Current;
                CT_RElt run = st.AddNewR();
                String fragment = text.Substring(runStartIdx, runEndIdx - runStartIdx);
                run.t = (fragment);
                PreserveSpaces(run.t);
                CT_RPrElt fmt = formats[runEndIdx];
                if (fmt != null)
                    run.rPr = (fmt);
                runStartIdx = runEndIdx;
            }
            return st;
        }

        private ThemesTable GetThemesTable()
        {
            if (styles == null) return null;
            return styles.GetTheme();
        }

        public bool HasFormatting()
        {
            //noinspection deprecation - for performance reasons!
            CT_RElt[] rs = st.r.ToArray();
            if (rs == null || rs.Length == 0)
            {
                return false;
            }
            foreach (CT_RElt r in rs)
            {
                //TODO: check that this functions the same.
                if (r.rPr != null) return true;
            }
            return false;
        }
    }
}

