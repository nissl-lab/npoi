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
    using System.Collections.Generic;
    using System;
    using System.Text;
    using NPOI.HWPF.SPRM;
    using System.IO;
    using NPOI.Util;
    using System.Diagnostics;
    /**
     * This class is the central class of the HWPF object model. All properties that
     * apply to a range of characters in a Word document extend this class.
     *
     * It is possible to insert text and/or properties at the beginning or end of a
     * range.
     *
     * Ranges are only valid if there hasn't been an insert in a prior Range since
     * the Range's creation. Once an element (text, paragraph, etc.) has been
     * inserted into a Range, subsequent Ranges become unstable.
     *
     * @author Ryan Ackley
     */
    public class Range : BaseObject
    { // TODO -instantiable superclass

        public const int TYPE_PARAGRAPH = 0;
        public const int TYPE_CHARACTER = 1;
        public const int TYPE_SECTION = 2;
        public const int TYPE_TEXT = 3;
        public const int TYPE_LISTENTRY = 4;
        public const int TYPE_TABLE = 5;
        public const int TYPE_UNDEFINED = 6;

        /** Needed so inserts and deletes will ripple up through Containing Ranges */
        internal Range _parent;

        /** The starting character offset of this range. */
        internal int _start;

        /** The ending character offset of this range. */
        internal int _end;

        /** The document this range blongs to. */
        protected HWPFDocumentCore _doc;

        /** Have we loaded the section indexes yet */
        bool _sectionRangeFound;

        /** All sections that belong to the document this Range belongs to. */
        protected List<SEPX> _sections;

        /** The start index in the sections list for this Range */
        protected int _sectionStart;

        /** The end index in the sections list for this Range. */
        protected int _sectionEnd;

        /** Have we loaded the paragraph indexes yet. */
        protected bool _parRangeFound;

        /** All paragraphs that belong to the document this Range belongs to. */
        internal List<PAPX> _paragraphs;

        /** The start index in the paragraphs list for this Range */
        internal int _parStart;

        /** The end index in the paragraphs list for this Range. */
        internal int _parEnd;

        /** Have we loaded the characterRun indexes yet. */
        protected bool _charRangeFound;

        /** All CharacterRuns that belong to the document this Range belongs to. */
        protected List<CHPX> _characters;

        /** The start index in the characterRuns list for this Range */
        protected int _charStart;

        /** The end index in the characterRuns list for this Range. */
        protected int _charEnd;

        /** Have we loaded the Text indexes yet */
        protected bool _textRangeFound;

        /** All text pieces that belong to the document this Range belongs to. */
        protected StringBuilder _text;

        /** The start index in the text list for this Range. */
        protected int _textStart;

        /** The end index in the text list for this Range. */
        protected int _textEnd;

        // protected Range()
        // {
        //
        // }

        /**
         * Used to construct a Range from a document. This is generally used to
         * create a Range that spans the whole document, or at least one whole part
         * of the document (eg main text, header, comment)
         *
         * @param start
         *            Starting character offset of the range.
         * @param end
         *            Ending character offset of the range.
         * @param doc
         *            The HWPFDocument the range is based on.
         */
        public Range(int start, int end, HWPFDocumentCore doc)
        {
            _start = start;
            _end = end;
            _doc = doc;
            _sections = _doc.SectionTable.GetSections();
            _paragraphs = _doc.ParagraphTable.GetParagraphs();
            _characters = _doc.CharacterTable.GetTextRuns();
            _text = _doc.Text;
            _parent = null;

            SanityCheckStartEnd();
        }


        /**
         * Used to create Ranges that are children of other Ranges.
         *
         * @param start
         *            Starting character offset of the range.
         * @param end
         *            Ending character offset of the range.
         * @param parent
         *            The parent this range belongs to.
         */
        internal Range(int start, int end, Range parent)
        {
            _start = start;
            _end = end;
            _doc = parent._doc;
            _sections = parent._sections;
            _paragraphs = parent._paragraphs;
            _characters = parent._characters;
            _text = parent._text;
            _parent = parent;

            SanityCheckStartEnd();
            SanityCheck();
        }


        /**
         * Ensures that the start and end were were given are actually valid, to
         * avoid issues later on if they're not
         */
        private void SanityCheckStartEnd()
        {
            if (_start < 0)
            {
                throw new ArgumentException("Range start must not be negative. Given " + _start);
            }
            if (_end < _start)
            {
                throw new ArgumentException("The end (" + _end
                        + ") must not be before the start (" + _start + ")");
            }
        }

        /**
         * Does any <code>TextPiece</code> in this Range use unicode?
         *
         * @return true if it does and false if it doesn't
         */
        [Obsolete]
        public bool UsesUnicode
        {

            get
            {
                return true;
            }
        }
        /**
         * Gets the text that this Range contains.
         *
         * @return The text for this range.
         */
        public String Text
        {
            get
            {
                return _text.ToString().Substring(_start, _end - _start);
            }
        }

        /**
         * Removes any fields (eg macros, page markers etc) from the string.
         * Normally used to make some text suitable for Showing to humans, and the
         * resultant text should not normally be saved back into the document!
         */
        public static String StripFields(String text)
        {
            // First up, fields can be nested...
            // A field can be 0x13 [contents] 0x15
            // Or it can be 0x13 [contents] 0x14 [real text] 0x15

            // If there are no fields, all easy
            if (text.IndexOf('\u0013') == -1)
                return text;

            // Loop over until they're all gone
            // That's when we're out of both 0x13s and 0x15s
            while (text.IndexOf('\u0013') > -1 && text.IndexOf('\u0015') > -1)
            {
                int first13 = text.IndexOf('\u0013');
                int next13 = text.IndexOf('\u0013', first13 + 1);
                int first14 = text.IndexOf('\u0014', first13 + 1);
                int last15 = text.LastIndexOf('\u0015');

                // If they're the wrong way around, give up
                if (last15 < first13)
                {
                    break;
                }

                // If no more 13s and 14s, just zap
                if (next13 == -1 && first14 == -1)
                {
                    text = text.Substring(0, first13) + text.Substring(last15 + 1);
                    break;
                }

                // If a 14 comes before the next 13, then
                // zap from the 13 to the 14, and remove
                // the 15
                if (first14 != -1 && (first14 < next13 || next13 == -1))
                {
                    text = text.Substring(0, first13) + text.Substring(first14 + 1, last15 - first14 - 1)
                            + text.Substring(last15 + 1);
                    continue;
                }

                // Another 13 comes before the next 14.
                // This means there's nested stuff, so we
                // can just zap the lot
                text = text.Substring(0, first13) + text.Substring(last15 + 1);
                continue;
            }

            return text;
        }

        /**
         * Used to get the number of sections in a range. If this range is smaller
         * than a section, it will return 1 for its Containing section.
         *
         * @return The number of sections in this range.
         */
        public int NumSections
        {
            get
            {
                InitSections();
                return _sectionEnd - _sectionStart;
            }
        }

        /**
         * Used to get the number of paragraphs in a range. If this range is smaller
         * than a paragraph, it will return 1 for its Containing paragraph.
         *
         * @return The number of paragraphs in this range.
         */

        public int NumParagraphs
        {
            get
            {
                InitParagraphs();
                return _parEnd - _parStart;
            }
        }

        /**
         *
         * @return The number of characterRuns in this range.
         */

        public int NumCharacterRuns
        {
            get
            {
                InitCharacterRuns();
                return _charEnd - _charStart;
            }
        }

        /**
         * Inserts text into the front of this range.
         *
         * @param text
         *            The text to insert
         * @return The character run that text was inserted into.
         */
        public CharacterRun InsertBefore(String text)
        // 
        {
            InitAll();

            _text.Insert(_start, text);

            _doc.CharacterTable.AdjustForInsert(_charStart, text.Length);
            _doc.ParagraphTable.AdjustForInsert(_parStart, text.Length);
            _doc.SectionTable.AdjustForInsert(_sectionStart, text.Length);
            AdjustForInsert(text.Length);

            // update the FIB.CCPText + friends fields
            AdjustFIB(text.Length);

            SanityCheck();

            return GetCharacterRun(0);
        }

        /**
         * Inserts text onto the end of this range
         *
         * @param text
         *            The text to insert
         * @return The character run the text was inserted into.
         */
        public CharacterRun InsertAfter(String text)
        {
            InitAll();

            _text.Insert(_end, text);
            _doc.CharacterTable.AdjustForInsert(_charEnd - 1, text.Length);
            _doc.ParagraphTable.AdjustForInsert(_parEnd - 1, text.Length);
            _doc.SectionTable.AdjustForInsert(_sectionEnd - 1, text.Length);
            AdjustForInsert(text.Length);

            return GetCharacterRun(NumCharacterRuns - 1);

        }

        /**
         * Inserts text into the front of this range and it gives that text the
         * CharacterProperties specified in props.
         *
         * @param text
         *            The text to insert.
         * @param props
         *            The CharacterProperties to give the text.
         * @return A new CharacterRun that has the given text and properties and is
         *         n ow a part of the document.
         */
        [Obsolete]
        public CharacterRun InsertBefore(String text, CharacterProperties props)
        // 
        {
            InitAll();
            PAPX papx = _paragraphs[_parStart];
            short istd = papx.GetIstd();

            StyleSheet ss = _doc.GetStyleSheet();
            CharacterProperties baseStyle = ss.GetCharacterStyle(istd);
            byte[] grpprl = CharacterSprmCompressor.CompressCharacterProperty(props, baseStyle);
            SprmBuffer buf = new SprmBuffer(grpprl);
            _doc.CharacterTable.Insert(_charStart, _start, buf);

            return InsertBefore(text);
        }

        /**
         * Inserts text onto the end of this range and gives that text the
         * CharacterProperties specified in props.
         *
         * @param text
         *            The text to insert.
         * @param props
         *            The CharacterProperties to give the text.
         * @return A new CharacterRun that has the given text and properties and is
         *         n ow a part of the document.
         */
        [Obsolete]
        public CharacterRun InsertAfter(String text, CharacterProperties props)
        // 
        {
            InitAll();
            PAPX papx = _paragraphs[_parEnd - 1];
            short istd = papx.GetIstd();

            StyleSheet ss = _doc.GetStyleSheet();
            CharacterProperties baseStyle = ss.GetCharacterStyle(istd);
            byte[] grpprl = CharacterSprmCompressor.CompressCharacterProperty(props, baseStyle);
            SprmBuffer buf = new SprmBuffer(grpprl);
            _doc.CharacterTable.Insert(_charEnd, _end, buf);
            _charEnd++;
            return InsertAfter(text);
        }

        /**
         * Inserts and empty paragraph into the front of this range.
         *
         * @param props
         *            The properties that the new paragraph will have.
         * @param styleIndex
         *            The index into the stylesheet for the new paragraph.
         * @return The newly inserted paragraph.
         */
        [Obsolete]
        public Paragraph InsertBefore(ParagraphProperties props, int styleIndex)
        // 
        {
            return this.InsertBefore(props, styleIndex, "\r");
        }

        /**
         * Inserts a paragraph into the front of this range. The paragraph will
         * contain one character run that has the default properties for the
         * paragraph's style.
         *
         * It is necessary for the text to end with the character '\r'
         *
         * @param props
         *            The paragraph's properties.
         * @param styleIndex
         *            The index of the paragraph's style in the style sheet.
         * @param text
         *            The text to insert.
         * @return A newly inserted paragraph.
         */
        [Obsolete]
        protected Paragraph InsertBefore(ParagraphProperties props, int styleIndex, String text)
        // 
        {
            InitAll();
            StyleSheet ss = _doc.GetStyleSheet();
            ParagraphProperties baseStyle = ss.GetParagraphStyle(styleIndex);
            CharacterProperties baseChp = ss.GetCharacterStyle(styleIndex);

            byte[] grpprl = ParagraphSprmCompressor.CompressParagraphProperty(props, baseStyle);
            byte[] withIndex = new byte[grpprl.Length + LittleEndianConsts.SHORT_SIZE];
            LittleEndian.PutShort(withIndex, (short)styleIndex);
            Array.Copy(grpprl, 0, withIndex, LittleEndianConsts.SHORT_SIZE, grpprl.Length);
            SprmBuffer buf = new SprmBuffer(withIndex);

            _doc.ParagraphTable.Insert(_parStart, _start, buf);
            InsertBefore(text, baseChp);
            return GetParagraph(0);
        }

        /**
         * Inserts and empty paragraph into the end of this range.
         *
         * @param props
         *            The properties that the new paragraph will have.
         * @param styleIndex
         *            The index into the stylesheet for the new paragraph.
         * @return The newly inserted paragraph.
         */
        [Obsolete]
        public Paragraph InsertAfter(ParagraphProperties props, int styleIndex)
        // 
        {
            return this.InsertAfter(props, styleIndex, "\r");
        }

        /**
         * Inserts a paragraph into the end of this range. The paragraph will
         * contain one character run that has the default properties for the
         * paragraph's style.
         *
         * It is necessary for the text to end with the character '\r'
         *
         * @param props
         *            The paragraph's properties.
         * @param styleIndex
         *            The index of the paragraph's style in the style sheet.
         * @param text
         *            The text to insert.
         * @return A newly inserted paragraph.
         */
        [Obsolete]
        protected Paragraph InsertAfter(ParagraphProperties props, int styleIndex, String text)
        // 
        {
            InitAll();
            StyleSheet ss = _doc.GetStyleSheet();
            ParagraphProperties baseStyle = ss.GetParagraphStyle(styleIndex);
            CharacterProperties baseChp = ss.GetCharacterStyle(styleIndex);

            byte[] grpprl = ParagraphSprmCompressor.CompressParagraphProperty(props, baseStyle);
            byte[] withIndex = new byte[grpprl.Length + LittleEndianConsts.SHORT_SIZE];
            LittleEndian.PutShort(withIndex, (short)styleIndex);
            Array.Copy(grpprl, 0, withIndex, LittleEndianConsts.SHORT_SIZE, grpprl.Length);
            SprmBuffer buf = new SprmBuffer(withIndex);

            _doc.ParagraphTable.Insert(_parEnd, _end, buf);
            _parEnd++;
            InsertAfter(text, baseChp);
            return GetParagraph(NumParagraphs - 1);
        }

        public void Delete()
        {

            InitAll();

            int numSections = _sections.Count;
            int numRuns = _characters.Count;
            int numParagraphs = _paragraphs.Count;

            for (int x = _charStart; x < numRuns; x++)
            {
                CHPX chpx = _characters[x];
                chpx.AdjustForDelete(_start, _end - _start);
            }

            for (int x = _parStart; x < numParagraphs; x++)
            {
                PAPX papx = _paragraphs[x];
                // System.err.println("Paragraph " + x + " was " + papx.Start +
                // " -> " + papx.End);
                papx.AdjustForDelete(_start, _end - _start);
                // System.err.println("Paragraph " + x + " is now " +
                // papx.Start + " -> " + papx.End);
            }

            for (int x = _sectionStart; x < numSections; x++)
            {
                SEPX sepx = _sections[x];
                // System.err.println("Section " + x + " was " + sepx.Start +
                // " -> " + sepx.End);
                sepx.AdjustForDelete(_start, _end - _start);
                // System.err.println("Section " + x + " is now " + sepx.Start
                // + " -> " + sepx.End);
            }
            _text = _text.Remove(_start, _end - _start);
            Range parent = _parent;
            if (parent != null)
            {
                parent.AdjustForInsert(-(_end - _start));
            }
            // update the FIB.CCPText + friends field
            AdjustFIB(-(_end - _start));
        }

        /**
         * Inserts a simple table into the beginning of this range. The number of
         * columns is determined by the TableProperties passed into this function.
         *
         * @param props
         *            The table properties for the table.
         * @param rows
         *            The number of rows.
         * @return The empty Table that is now part of the document.
         */
        [Obsolete]
        public Table InsertBefore(TableProperties props, int rows)
        {
            ParagraphProperties parProps = new ParagraphProperties();
            parProps.SetFInTable(true);
            parProps.SetItap(1);

            int columns = props.GetItcMac();
            for (int x = 0; x < rows; x++)
            {
                Paragraph cell = this.InsertBefore(parProps, StyleSheet.NIL_STYLE);
                cell.InsertAfter("\u0007");
                for (int y = 1; y < columns; y++)
                {
                    cell = cell.InsertAfter(parProps, StyleSheet.NIL_STYLE);
                    cell.InsertAfter("\u0007");
                }
                cell = cell.InsertAfter(parProps, StyleSheet.NIL_STYLE, "\u0007");
                cell.SetTableRowEnd(props);
            }
            return new Table(_start, _start + (rows * (columns + 1)), this, 1);
        }

        /**
         * Inserts a simple table into the beginning of this range.
         * 
         * @param columns
         *            The number of columns
         * @param rows
         *            The number of rows.
         * @return The empty Table that is now part of the document.
         */
        public Table InsertTableBefore(short columns, int rows)
        {
            ParagraphProperties parProps = new ParagraphProperties();
            parProps.SetFInTable(true);
            parProps.SetItap(1);

            int oldEnd = this._end;

            for (int x = 0; x < rows; x++)
            {
                Paragraph cell = this.InsertBefore(parProps, StyleSheet.NIL_STYLE);
                cell.InsertAfter(Convert.ToString('\u0007'));
                for (int y = 1; y < columns; y++)
                {
                    cell = cell.InsertAfter(parProps, StyleSheet.NIL_STYLE);
                    cell.InsertAfter(Convert.ToString('\u0007'));
                }
                cell = cell.InsertAfter(parProps, StyleSheet.NIL_STYLE,
                        Convert.ToString('\u0007'));
                cell.SetTableRowEnd(new TableProperties(columns));
            }

            int newEnd = this._end;
            int diff = newEnd - oldEnd;

            return new Table(_start, _start + diff, this, 1);
        }
        /**
         * Inserts a list into the beginning of this range.
         *
         * @param props
         *            The properties of the list entry. All list entries are
         *            paragraphs.
         * @param listID
         *            The id of the list that Contains the properties.
         * @param level
         *            The indentation level of the list.
         * @param styleIndex
         *            The base style's index in the stylesheet.
         * @return The empty ListEntry that is now part of the document.
         */
        public ListEntry InsertBefore(ParagraphProperties props, int listID, int level, int styleIndex)
        {
            ListTables lt = _doc.GetListTables();
            if (lt.GetLevel(listID, level) == null)
            {
                throw new InvalidDataException("The specified list and level do not exist");
            }

            int ilfo = lt.GetOverrideIndexFromListID(listID);
            props.SetIlfo(ilfo);
            props.SetIlvl((byte)level);

            return (ListEntry)InsertBefore(props, styleIndex);
        }

        /**
         * Inserts a list into the beginning of this range.
         *
         * @param props
         *            The properties of the list entry. All list entries are
         *            paragraphs.
         * @param listID
         *            The id of the list that Contains the properties.
         * @param level
         *            The indentation level of the list.
         * @param styleIndex
         *            The base style's index in the stylesheet.
         * @return The empty ListEntry that is now part of the document.
         */
        public ListEntry InsertAfter(ParagraphProperties props, int listID, int level, int styleIndex)
        {
            ListTables lt = _doc.GetListTables();
            if (lt.GetLevel(listID, level) == null)
            {
                throw new InvalidDataException("The specified list and level do not exist");
            }
            int ilfo = lt.GetOverrideIndexFromListID(listID);
            props.SetIlfo(ilfo);
            props.SetIlvl((byte)level);

            return (ListEntry)InsertAfter(props, styleIndex);
        }

        /**
         * Replace (one instance of) a piece of text with another...
         *
         * @param pPlaceHolder
         *            The text to be Replaced (e.g., "${organization}")
         * @param pValue
         *            The Replacement text (e.g., "Apache Software Foundation")
         * @param pOffset
         *            The offset or index where the text to be Replaced begins
         *            (relative to/within this <code>Range</code>)
         */
        public void ReplaceText(String pPlaceHolder, String pValue, int pOffSet)
        {
            int absPlaceHolderIndex = StartOffset + pOffSet;
            Range subRange = new Range(absPlaceHolderIndex, (absPlaceHolderIndex + pPlaceHolder
                    .Length), this);

            subRange.InsertBefore(pValue);

            // re-create the sub-range so we can delete it
            subRange = new Range((absPlaceHolderIndex + pValue.Length), (absPlaceHolderIndex
                    + pPlaceHolder.Length + pValue.Length), GetDocument());

            // deletes are automagically propagated
            subRange.Delete();
        }

        /**
         * Replace (all instances of) a piece of text with another...
         *
         * @param pPlaceHolder
         *            The text to be Replaced (e.g., "${organization}")
         * @param pValue
         *            The Replacement text (e.g., "Apache Software Foundation")
         */
        public void ReplaceText(String pPlaceHolder, String pValue)
        {
            bool keepLooking = true;
            while (keepLooking)
            {

                String text = Text;
                int offset = text.IndexOf(pPlaceHolder);
                if (offset >= 0)
                    ReplaceText(pPlaceHolder, pValue, offset);
                else
                    keepLooking = false;
            }
        }

        /**
         * Gets the character run at index. The index is relative to this range.
         *
         * @param index
         *            The index of the character run to Get.
         * @return The character run at the specified index in this range.
         */
        public CharacterRun GetCharacterRun(int index)
        {
            InitCharacterRuns();

            if (index + _charStart >= _charEnd)
                throw new IndexOutOfRangeException("CHPX #" + index + " ("
                        + (index + _charStart) + ") not in range [" + _charStart
                        + "; " + _charEnd + ")");

            CHPX chpx = _characters[index + _charStart];

            if (chpx == null)
            {
                return null;
            }
                    short istd;
            if ( this is Paragraph )
            {
                istd = ((Paragraph) this)._istd;
            }
            else
            {
            int[] point = FindRange(_paragraphs, Math.Max(chpx.Start, _start), Math.Min(chpx.End, _end));
            InitParagraphs();
 			int parStart = Math.Max( point[0], _parStart );
            if ( parStart >= _paragraphs.Count )
            {
                return null;
            }
            PAPX papx = _paragraphs[point[0]];
            istd = papx.GetIstd();
			}

            CharacterRun chp = new CharacterRun(chpx, _doc.GetStyleSheet(), istd, this);

            return chp;
        }

        /**
         * Gets the section at index. The index is relative to this range.
         *
         * @param index
         *            The index of the section to Get.
         * @return The section at the specified index in this range.
         */
        public Section GetSection(int index)
        {
            InitSections();
            SEPX sepx = _sections[index + _sectionStart];
            Section sep = new Section(sepx, this);
            return sep;
        }

        /**
         * Gets the paragraph at index. The index is relative to this range.
         *
         * @param index
         *            The index of the paragraph to Get.
         * @return The paragraph at the specified index in this range.
         */

        public Paragraph GetParagraph(int index)
        {
            InitParagraphs();
            PAPX papx = _paragraphs[index + _parStart];

            ParagraphProperties props = papx.GetParagraphProperties(_doc.GetStyleSheet());
            Paragraph pap = null;
            if (props.GetIlfo() > 0)
            {
                pap = new ListEntry(papx, this, _doc.GetListTables());
            }
            else
            {
                if (((index + _parStart) == 0) && papx.Start > 0)
                {
                    pap = new Paragraph(papx, this, 0);
                }
                else
                {
                    pap = new Paragraph(papx, this);
                }
            }

            return pap;
        }

        /**
         * This method is used to determine the type. Handy for switch statements
         * Compared to the is operator.
         *
         * @return A TYPE constant.
         */
        public virtual int Type
        {
            get
            {
                return TYPE_UNDEFINED;
            }
        }

        /**
         * Gets the table that starts with paragraph. In a Word file, a table
         * consists of a group of paragraphs with certain flags Set.
         *
         * @param paragraph
         *            The paragraph that is the first paragraph in the table.
         * @return The table that starts with paragraph
         */
        public Table GetTable(Paragraph paragraph)
        {
            if (!paragraph.IsInTable())
            {
                throw new ArgumentException("This paragraph doesn't belong to a table");
            }

            Range r = paragraph;
            if (r._parent != this)
            {
                throw new ArgumentException("This paragraph is not a child of this range");
            }

            r.InitAll();
            int tableLevel = paragraph.GetTableLevel();
            int tableEndInclusive = r._parStart;

            if (r._parStart != 0)
            {
                Paragraph previous = new Paragraph(
                   _paragraphs[r._parStart - 1], this);
                if (previous.IsInTable()
                    && previous.GetTableLevel() == tableLevel
                    && previous._sectionEnd >= r._sectionStart)
                    throw new ArgumentException("This paragraph is not the first one in the table");
            }

            Range overallRange = _doc.GetOverallRange();
            int limit = _paragraphs.Count;
            for (; tableEndInclusive < limit; tableEndInclusive++)
            {
                Paragraph next = new Paragraph(
                    _paragraphs[tableEndInclusive + 1], overallRange);
                if (!next.IsInTable() || next.GetTableLevel() < tableLevel)
                    break;
            }

            InitAll();
            if (tableEndInclusive > _parEnd)
            {
                throw new IndexOutOfRangeException(
                        "The table's bounds fall outside of this Range");
            }
            if (tableEndInclusive < 0)
            {
                throw new IndexOutOfRangeException(
                        "The table's end is negative, which isn't allowed!");
            }
            int endOffsetExclusive = _paragraphs[tableEndInclusive].End;

            return new Table(paragraph.StartOffset, endOffsetExclusive, this, paragraph.GetTableLevel());
        }

        /**
         * loads all of the list indexes.
         */
        protected void InitAll()
        {
            InitCharacterRuns();
            InitParagraphs();
            InitSections();
        }

        /**
         * Inits the paragraph list indexes.
         */
        private void InitParagraphs()
        {
            if (!_parRangeFound)
            {
                int[] point = FindRange<PAPX>(_paragraphs, _start, _end);
                _parStart = point[0];
                _parEnd = point[1];
                _parRangeFound = true;
            }
        }

        /**
         * Inits the character run list indexes.
         */
        private void InitCharacterRuns()
        {
            if (!_charRangeFound)
            {
                int[] point = FindRange<CHPX>(_characters, _start, _end);
                _charStart = point[0];
                _charEnd = point[1];
                _charRangeFound = true;
            }
        }

        /**
         * Inits the section list indexes.
         */
        private void InitSections()
        {
            if (!_sectionRangeFound)
            {
                int[] point = FindRange<SEPX>(_sections,_sectionStart, _start, _end);
                _sectionStart = point[0];
                _sectionEnd = point[1];
                _sectionRangeFound = true;
            }
        }


        private static int BinarySearchStart<T>(List<T> rpl,
            int start) where T : PropertyNode
        {
            if (rpl[0].Start >= start)
                return 0;

            int low = 0;
            int high = rpl.Count - 1;

            while (low <= high)
            {
                int mid = (low + high) >> 1;
                PropertyNode node = rpl[mid];

                if (node.Start < start)
                {
                    low = mid + 1;
                }
                else if (node.Start > start)
                {
                    high = mid - 1;
                }
                else
                {
                    return mid;
                }
            }
            return low - 1;
        }
        private static int BinarySearchEnd<T>(List<T> rpl,
                int foundStart, int end) where T : PropertyNode
        {
            if (rpl[rpl.Count - 1].End <= end)
                return rpl.Count - 1;

            int low = foundStart;
            int high = rpl.Count - 1;

            while (low <= high)
            {
                int mid = (low + high) >> 1;
                PropertyNode node = rpl[mid];

                if (node.End < end)
                {
                    low = mid + 1;
                }
                else if (node.End > end)
                {
                    high = mid - 1;
                }
                else
                {
                    return mid;
                }
            }
            return low;
        }



        /**
         * Used to find the list indexes of a particular property.
         *
         * @param rpl
         *            A list of property nodes.
         * @param min
         *            A hint on where to start looking.
         * @param start
         *            The starting character offSet.
         * @param end
         *            The ending character offSet.
         * @return An int array of length 2. The first int is the start index and
         *         the second int is the end index.
         */
        private int[] FindRange<T>(List<T> rpl, int start, int end) where T : PropertyNode
        {
            int startIndex = BinarySearchStart(rpl, start);
            while (startIndex > 0 && rpl[startIndex - 1].Start >= start)
                startIndex--;

            int endIndex = BinarySearchEnd(rpl, startIndex, end);
            while (endIndex < rpl.Count - 1
                    && rpl[endIndex + 1].End <= end)
                endIndex++; 

            if (startIndex < 0 || startIndex >= rpl.Count
                    || startIndex > endIndex || endIndex < 0
                    || endIndex >= rpl.Count)
                throw new Exception();

            return new int[] { startIndex, endIndex + 1 };
        }
	/**
	 * Used to find the list indexes of a particular property.
	 *
	 * @param rpl
	 *            A list of property nodes.
	 * @param min
	 *            A hint on where to start looking.
	 * @param start
	 *            The starting character offset.
	 * @param end
	 *            The ending character offset.
	 * @return An int array of length 2. The first int is the start index and
	 *         the second int is the end index.
	 */
	private int[] FindRange<T>(List<T> rpl, int min, int start, int end) where T : PropertyNode 
    {
		int x = min;
		
        if ( rpl.Count == min )
            return new int[] { min, min };

        T node = rpl[x];

		while (node==null || (node.End <= start && x < rpl.Count - 1)) {
			x++;

            if (x>=rpl.Count) {
                return new int[] {0, 0};
            }

			node = rpl[x];
		}

        if ( node.Start > end )
        {
            return new int[] { 0, 0 };
        }

        if ( node.End <= start )
        {
            return new int[] { rpl.Count, rpl.Count };
        }

        for ( int y = x; y < rpl.Count; y++ )
        {
            node = rpl[y];
            if ( node == null )
                continue;

            if ( node.Start < end && node.End <= end )
                continue;

            if ( node.Start < end )
                return new int[] { x, y +1 };

            return new int[] { x, y };
        }
        return new int[] { x, rpl.Count };
    }
        /**
         * reSets the list indexes.
         */
        protected virtual void Reset()
        {
            _textRangeFound = false;
            _charRangeFound = false;
            _parRangeFound = false;
            _sectionRangeFound = false;
        }

        /**
         * Adjust the value of the various FIB character count fields, eg
         * <code>FIB.CCPText</code> after an insert or a delete...
         *
         * Works on all CCP fields from this range onwards
         *
         * @param adjustment
         *            The (signed) value that should be Added to the FIB CCP fields
         */
        protected void AdjustFIB(int adjustment)
        {
            //Assert(_doc is HWPFDocument);

            // update the FIB.CCPText field (this should happen once per adjustment,
            // so we don't want it in
            // adjustForInsert() or it would get updated multiple times if the range
            // has a parent)
            // without this, OpenOffice.org (v. 2.2.x) does not see all the text in
            // the document

            //CPSplitCalculator cpS = ((HWPFDocument)_doc).GetCPSplitCalculator();
            FileInformationBlock fib = _doc.GetFileInformationBlock();

            // Do for each affected part
            //if (_start < cpS.GetMainDocumentEnd())
            //{
            //    fib.SetCcpText(fib.GetCcpText() + adjustment);
            //}

            //if (_start < cpS.GetCommentsEnd())
            //{
            //    fib.SetCcpAtn(fib.GetCcpAtn() + adjustment);
            //}
            //if (_start < cpS.GetEndNoteEnd())
            //{
            //    fib.SetCcpEdn(fib.GetCcpEdn() + adjustment);
            //}
            //if (_start < cpS.GetFootnoteEnd())
            //{
            //    fib.SetCcpFtn(fib.GetCcpFtn() + adjustment);
            //}
            //if (_start < cpS.GetHeaderStoryEnd())
            //{
            //    fib.SetCcpHdd(fib.GetCcpHdd() + adjustment);
            //}
            //if (_start < cpS.GetHeaderTextboxEnd())
            //{
            //    fib.SetCcpHdrTxtBx(fib.GetCcpHdrTxtBx() + adjustment);
            //}
            //if (_start < cpS.GetMainTextboxEnd())
            //{
            //    fib.SetCcpTxtBx(fib.GetCcpTxtBx() + adjustment);
            //}

            int currentEnd = 0;
            foreach (SubdocumentType type in HWPFDocument.ORDERED)
            {
                int currentLength = fib.GetSubdocumentTextStreamLength(type);
                currentEnd += currentLength;

                // do we need to shift this part?
                if (_start > currentEnd)
                    continue;

                fib.SetSubdocumentTextStreamLength(type, currentLength
                        + adjustment);

                break;
            }
        }

        /**
         * adjust this range after an insert happens.
         *
         * @param length
         *            the length to adjust for (expected to be a count of
         *            code-points, not necessarily chars)
         */
        private void AdjustForInsert(int length)
        {
            _end += length;

            Reset();
            Range parent = (Range)_parent;
            if (parent != null)
            {
                parent.AdjustForInsert(length);
            }
        }


        public int StartOffset
        {
            get
            {
                return _start;
            }
        }

        public int EndOffset
        {
            get
            {
                return _end;
            }
        }



        internal HWPFDocumentCore GetDocument()
        {
            return _doc;
        }


        public override String ToString()
        {
            return "Range from " + StartOffset + " to " + EndOffset
                    + " (chars)";
        }
        /**
          * Method for debug purposes. Checks that all resolved elements are inside
          * of current range.
          */
        public bool SanityCheck()
        {
            Debug.Assert(_start >= 0, "start < 0");
            Debug.Assert(_start <= _text.Length, "start > text length");
            Debug.Assert(_end >= 0, "end < 0");
            Debug.Assert(_end <= _text.Length, "end > text length");
            Debug.Assert(_start <= _end, "start > end");


            if (_charRangeFound)
            {
                for (int c = _charStart; c < _charEnd; c++)
                {
                    CHPX chpx = _characters[c];

                    int left = Math.Max(this._start, chpx.Start);
                    int right = Math.Min(this._end, chpx.End);
                    Debug.Assert(left < right, "left >= right");
                }
            }
            if (_parRangeFound)
            {
                for (int p = _parStart; p < _parEnd; p++)
                {
                    PAPX papx = _paragraphs[p];

                    int left = Math.Max(this._start, papx.Start);
                    int right = Math.Min(this._end, papx.End);

                    Debug.Assert(left < right, "left >= right");
                }
            }

            return true;
        }
    }
}