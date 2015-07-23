/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Collections;

    using NPOI.HSSF.Record;
    using NPOI.HSSF.Model;
    using System.Collections.Generic;

    /// <summary>
    /// Rich text Unicode string.  These strings can have fonts applied to
    /// arbitary parts of the string.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// @author Jason Height (jheight at apache.org)
    /// </summary> 
    [Serializable]
    public class HSSFRichTextString : IComparable<HSSFRichTextString>,NPOI.SS.UserModel.IRichTextString
    {
        /** Place holder for indicating that NO_FONT has been applied here */
        public const short NO_FONT = 0;

        [NonSerialized]
        private UnicodeString _string;
        private InternalWorkbook _book;
        private LabelSSTRecord _record;

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFRichTextString"/> class.
        /// </summary>
        public HSSFRichTextString()
            : this("")
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFRichTextString"/> class.
        /// </summary>
        /// <param name="str">The string.</param>
        public HSSFRichTextString(String str)
        {
            if (str == null)
            {
                _string = new UnicodeString("");
            }
            else
            {
                this._string = new UnicodeString(str);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFRichTextString"/> class.
        /// </summary>
        /// <param name="book">The workbook.</param>
        /// <param name="record">The record.</param>
        public HSSFRichTextString(InternalWorkbook book, LabelSSTRecord record)
        {
            SetWorkbookReferences(book, record);

            this._string = book.GetSSTString(record.SSTIndex);
        }

        /// <summary>
        /// This must be called to Setup the internal work book references whenever
        /// a RichTextString Is Added to a cell
        /// </summary>
        /// <param name="book">The workbook.</param>
        /// <param name="record">The record.</param>
        public void SetWorkbookReferences(InternalWorkbook book, LabelSSTRecord record)
        {
            this._book = book;
            this._record = record;
        }

        /// <summary>
        /// Called whenever the Unicode string Is modified. When it Is modified
        /// we need to Create a new SST index, so that other LabelSSTRecords will not
        /// be affected by Changes tat we make to this string.
        /// </summary>
        /// <returns></returns>
        private UnicodeString CloneStringIfRequired()
        {
            if (_book == null)
                return _string;
            UnicodeString s = (UnicodeString)_string.Clone();
            return s;
        }

        /// <summary>
        /// Adds to SST if required.
        /// </summary>
        private void AddToSSTIfRequired()
        {
            if (_book != null)
            {
                int index = _book.AddSSTString(_string);
                _record.SSTIndex = (index);
                //The act of Adding the string to the SST record may have meant that
                //a extsing string was returned for the index, so update our local version
                _string = _book.GetSSTString(index);
            }
        }


        /// <summary>
        /// Applies a font to the specified Chars of a string.
        /// </summary>
        /// <param name="startIndex">The start index to apply the font to (inclusive).</param>
        /// <param name="endIndex">The end index to apply the font to (exclusive).</param>
        /// <param name="fontIndex">The font to use.</param>
        public void ApplyFont(int startIndex, int endIndex, short fontIndex)
        {
            if (startIndex > endIndex)
                throw new ArgumentException("Start index must be less than end index.");
            if (startIndex < 0 || endIndex > Length)
                throw new ArgumentException("Start and end index not in range.");
            if (startIndex == endIndex)
                return;

            //Need to Check what the font Is currently, so we can reapply it after
            //the range Is completed
            short currentFont = NO_FONT;
            if (endIndex != Length)
            {
                currentFont = this.GetFontAtIndex(endIndex);
            }

            //Need to clear the current formatting between the startIndex and endIndex
            _string = CloneStringIfRequired();
            System.Collections.Generic.List<UnicodeString.FormatRun> formatting = _string.FormatIterator();

            ArrayList deletedFR = new ArrayList();
            if (formatting != null)
            {
                IEnumerator<UnicodeString.FormatRun> formats = formatting.GetEnumerator();
                while (formats.MoveNext())
                {
                    UnicodeString.FormatRun r = formats.Current;
                    if ((r.CharacterPos >= startIndex) && (r.CharacterPos < endIndex))
                    {
                        deletedFR.Add(r);
                    }
                }
            }
            foreach (UnicodeString.FormatRun fr in deletedFR)
            {
                _string.RemoveFormatRun(fr);
            }

            _string.AddFormatRun(new UnicodeString.FormatRun((short)startIndex, fontIndex));
            if (endIndex != Length)
                _string.AddFormatRun(new UnicodeString.FormatRun((short)endIndex, currentFont));

            AddToSSTIfRequired();
        }

        /// <summary>
        /// Applies a font to the specified Chars of a string.
        /// </summary>
        /// <param name="startIndex">The start index to apply the font to (inclusive).</param>
        /// <param name="endIndex"> The end index to apply to font to (exclusive).</param>
        /// <param name="font">The index of the font to use.</param>
        public void ApplyFont(int startIndex, int endIndex, NPOI.SS.UserModel.IFont font)
        {
            ApplyFont(startIndex, endIndex, font.Index);
        }

        /// <summary>
        /// Sets the font of the entire string.
        /// </summary>
        /// <param name="font">The font to use.</param>
        public void ApplyFont(NPOI.SS.UserModel.IFont font)
        {
            ApplyFont(0, _string.CharCount, font);
        }

        /// <summary>
        /// Removes any formatting that may have been applied to the string.
        /// </summary>
        public void ClearFormatting()
        {
            _string = CloneStringIfRequired();
            _string.ClearFormatting();
            AddToSSTIfRequired();
        }

        /// <summary>
        /// Returns the plain string representation.
        /// </summary>
        /// <value>The string.</value>
        public String String
        {
            get { return _string.String; }
        }
        /// <summary>
        /// Returns the raw, probably shared Unicode String.
        /// Used when tweaking the styles, eg updating font
        /// positions.
        /// Changes to this string may well effect
        /// other RichTextStrings too!
        /// </summary>
        /// <value>The raw unicode string.</value>
        public UnicodeString RawUnicodeString
        {
            get
            {
                return _string;
            }
        }

        /// <summary>
        /// Gets or sets the unicode string.
        /// </summary>
        /// <value>The unicode string.</value>
        public UnicodeString UnicodeString
        {
            get { return CloneStringIfRequired(); }
            set { this._string = value; }
        }


        /// <summary>
        /// Gets the number of Chars in the font..
        /// </summary>
        /// <value>The length.</value>
        public int Length
        {
            get { return _string.CharCount; }
        }

        /// <summary>
        /// Returns the font in use at a particular index.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns>The font that's currently being applied at that
        /// index or null if no font Is being applied or the
        /// index Is out of range.</returns>
        public short GetFontAtIndex(int index)
        {
            int size = _string.FormatRunCount;
            UnicodeString.FormatRun currentRun = null;
            for (int i = 0; i < size; i++)
            {
                UnicodeString.FormatRun r = _string.GetFormatRun(i);
                if (r.CharacterPos > index)
                    break;
                else currentRun = r;
            }
            if (currentRun == null)
                return NO_FONT;
            else return currentRun.FontIndex;
        }

        /// <summary>
        /// Gets the number of formatting runs used. There will always be at
        /// least one of font NO_FONT.
        /// </summary>
        /// <value>The num formatting runs.</value>
        public int NumFormattingRuns
        {
            get { return _string.FormatRunCount; }
        }

        /// <summary>
        /// The index within the string to which the specified formatting run applies.
        /// </summary>
        /// <param name="index">the index of the formatting run</param>
        /// <returns>the index within the string.</returns>
        public int GetIndexOfFormattingRun(int index)
        {
            UnicodeString.FormatRun r = _string.GetFormatRun(index);
            return r.CharacterPos;
        }

        /// <summary>
        /// Gets the font used in a particular formatting run.
        /// </summary>
        /// <param name="index">the index of the formatting run.</param>
        /// <returns>the font number used.</returns>
        public short GetFontOfFormattingRun(int index)
        {
            UnicodeString.FormatRun r = _string.GetFormatRun(index);
            return r.FontIndex;
        }

        /// <summary>
        /// Compares one rich text string to another.
        /// </summary>
        /// <param name="other">The other rich text string.</param>
        /// <returns></returns>
        public int CompareTo(HSSFRichTextString other)
        {
            return _string.CompareTo(other._string);
        }

        /// <summary>
        /// Equalses the specified o.
        /// </summary>
        /// <param name="o">The o.</param>
        /// <returns></returns>
        public override bool Equals(Object o)
        {
            if (o is HSSFRichTextString)
            {
                return _string.Equals(((HSSFRichTextString)o)._string);
            }
            return false;
        }

        public override int GetHashCode ()
        {
            return _string.GetHashCode ();
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            return _string.ToString();
        }

        /// <summary>
        /// Applies the specified font to the entire string.
        /// </summary>
        /// <param name="fontIndex">Index of the font to apply.</param>
        public void ApplyFont(short fontIndex)
        {
            ApplyFont(0, _string.CharCount, fontIndex);
        }
    }
}