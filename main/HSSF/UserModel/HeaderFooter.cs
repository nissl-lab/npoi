/* ====================================================================
   Licensed to the Apache SoftwAre Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, softwAre
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Collections;
    using System.Text.RegularExpressions;
    using System.Text;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Common class for HSSFHeader and HSSFFooter
    /// </summary>
    public abstract class HeaderFooter:NPOI.SS.UserModel.IHeaderFooter
    {
        protected bool stripFields = false;

        /**
         * @return the internal text representation (combining center, left and right parts).
         * Possibly empty string if no header or footer is set.  Never <c>null</c>.
         */
        public abstract String RawText { get; }
        
        private String[] SplitParts()
        {
            String text = RawText;
            // default values
            String _left = "";
            String _center = "";
            String _right = "";

            while (text.Length > 1)
            {
                if (text[0] != '&') 
                {
                        _center = text;
                        break;
                }
                int pos = text.Length;
                switch (text[1])
                {
                    case 'L':
                        if (text.IndexOf("&C", StringComparison.Ordinal) >= 0)
                        {
                            pos = Math.Min(pos, text.IndexOf("&C", StringComparison.Ordinal));
                        }
                        if (text.IndexOf("&R", StringComparison.Ordinal) >= 0)
                        {
                            pos = Math.Min(pos, text.IndexOf("&R", StringComparison.Ordinal));
                        }
                        _left = text.Substring(2, pos - 2);
                        text = text.Substring(pos);
                        break;
                    case 'C':
                        if (text.IndexOf("&L", StringComparison.Ordinal) >= 0)
                        {
                            pos = Math.Min(pos, text.IndexOf("&L", StringComparison.Ordinal));
                        }
                        if (text.IndexOf("&R", StringComparison.Ordinal) >= 0)
                        {
                            pos = Math.Min(pos, text.IndexOf("&R", StringComparison.Ordinal));
                        }
                        _center = text.Substring(2, pos - 2);
                        text = text.Substring(pos);
                        break;
                    case 'R':
                        if (text.IndexOf("&C", StringComparison.Ordinal) >= 0)
                        {
                            pos = Math.Min(pos, text.IndexOf("&C", StringComparison.Ordinal));
                        }
                        if (text.IndexOf("&L", StringComparison.Ordinal) >= 0)
                        {
                            pos = Math.Min(pos, text.IndexOf("&L", StringComparison.Ordinal));
                        }
                        _right = text.Substring(2, pos - 2);
                        text = text.Substring(pos);
                        break;
                    default:
                        _center = text;
                        break;
                }
            }
            return new String[] { _left, _center, _right, };
        }
        private void UpdatePart(int partIndex, String newValue)
        {
            String[] parts = SplitParts();
            parts[partIndex] = newValue == null ? "" : newValue;
            UpdateHeaderFooterText(parts);
        }
        /// <summary>
        /// Creates the complete footer string based on the left, center, and middle
        /// strings.
        /// </summary>
        /// <param name="parts">The parts.</param>
        private void UpdateHeaderFooterText(String[] parts)
        {
            String _left = parts[0];
            String _center = parts[1];
            String _right = parts[2];

            if (_center.Length < 1 && _left.Length < 1 && _right.Length < 1)
            {
                SetHeaderFooterText(string.Empty);
                return;
            }
            StringBuilder sb = new StringBuilder(64);
            sb.Append("&C");
            sb.Append(_center);
            sb.Append("&L");
            sb.Append(_left);
            sb.Append("&R");
            sb.Append(_right);
            String text = sb.ToString();
            SetHeaderFooterText(text);
        }
        protected HeaderFooter()
        {

        }
        /// <summary>
        /// Sets the header footer text.
        /// </summary>
        /// <param name="text">the new header footer text (contains mark-up tags). Possibly
        /// empty string never </param>
        protected abstract void SetHeaderFooterText(String text);

        /// <summary>
        /// Get the left side of the header or footer.
        /// </summary>
        /// <value>The string representing the left side.</value>
        public String Left 
        {
            get
            {
                return SplitParts()[0];
            }
            set 
            {
                UpdatePart(0, value); 
            }
        }
        /// <summary>
        /// Get the center of the header or footer.
        /// </summary>
        /// <value>The string representing the center.</value>
        public String Center
        {
            get
            {
                return SplitParts()[1];
            }
            set
            {
                UpdatePart(1, value);
            }
        }

        /// <summary>
        /// Get the right side of the header or footer.
        /// </summary>
        /// <value>The string representing the right side..</value>
        public String Right
        {
            get
            {
                return SplitParts()[2];
            }
            set
            {
                UpdatePart(2, value);
            }
        }


        /// <summary>
        /// Returns the string that represents the change in font size.
        /// </summary>
        /// <param name="size">the new font size.</param>
        /// <returns>The special string to represent a new font size</returns>
        public static String FontSize(short size)
        {
            return "&" + size;
        }

        /// <summary>
        /// Returns the string that represents the change in font.
        /// </summary>
        /// <param name="font">the new font.</param>
        /// <param name="style">the fonts style, one of regular, italic, bold, italic bold or bold italic.</param>
        /// <returns>The special string to represent a new font size</returns>
        public static String Font(String font, String style)
        {
            return "&\"" + font + "," + style + "\"";
        }

        /// <summary>
        /// Returns the string representing the current page number
        /// </summary>
        /// <value>The special string for page number.</value>
        public static String Page
        {
            get
            {
                return PAGE_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the number of pages.
        /// </summary>
        /// <value>The special string for the number of pages.</value>
        public static String NumPages
        {
            get
            {
                return NUM_PAGES_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the current date
        /// </summary>
        /// <value>The special string for the date</value>
        public static String Date
        {
            get
            {
                return DATE_FIELD.sequence;
            }
        }

        /// <summary>
        /// Gets the time.
        /// </summary>
        /// <value>The time.</value>
        /// Returns the string representing the current time
        /// @return The special string for the time
        public static String Time
        {
            get
            {
                return TIME_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the current file name
        /// </summary>
        /// <value>The special string for the file name.</value>
        public static String File
        {
            get{
                return FILE_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the current tab (sheet) name
        /// </summary>
        /// <value>The special string for tab name.</value>
        public static String Tab
        {
            get
            {
                return SHEET_NAME_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the start bold
        /// </summary>
        /// <returns>The special string for start bold</returns>
        public static String StartBold
        {
            get
            {
                return BOLD_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the end bold
        /// </summary>
        /// <value>The special string for end bold.</value>
        public static String EndBold
        {
            get
            {
                return BOLD_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the start underline
        /// </summary>
        /// <value>The special string for start underline.</value>
        public static String StartUnderline
        {
            get
            {
                return UNDERLINE_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the end underline
        /// </summary>
        /// <value>The special string for end underline.</value>
        public static String EndUnderline
        {
            get
            {
                return UNDERLINE_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the start double underline
        /// </summary>
        /// <value>The special string for start double underline.</value>
        public static String StartDoubleUnderline
        {
            get
            {
                return DOUBLE_UNDERLINE_FIELD.sequence;
            }
        }

        /// <summary>
        /// Returns the string representing the end double underline
        /// </summary>
        /// <value>The special string for end double underline.</value>
        public static String EndDoubleUnderline
        {
            get
            {
                return DOUBLE_UNDERLINE_FIELD.sequence;
            }
        }


        /// <summary>
        /// Removes any fields (eg macros, page markers etc)
        /// from the string.
        /// Normally used to make some text suitable for showing
        /// to humans, and the resultant text should not normally
        /// be saved back into the document!
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        public static String StripFields(String text)
        {
            int pos;

            // Check we really got something to work on
            if (text == null || text.Length == 0)
            {
                return text;
            }

            foreach(Field field in Fields.AllFields)
            {
                String seq = field.sequence;
                while ((pos = text.IndexOf(seq, StringComparison.CurrentCulture)) > -1)
                {
                    text = text.Substring(0, pos) +
                        text.Substring(pos + seq.Length);
                }
            }

            // Now do the tricky, dynamic ones
            // These are things like font sizes and font names

            text = Regex.Replace(text,@"\&\d+", "");
            text = Regex.Replace(text,"\\&\".*?,.*?\"", "");

            // All done
            return text;
        }


        /// <summary>
        /// Are fields currently being Stripped from
        /// the text that this {@link HeaderStories} returns?
        /// Default is false, but can be changed
        /// </summary>
        /// <value><c>true</c> if [are fields stripped]; otherwise, <c>false</c>.</value>
        public bool AreFieldsStripped
        {
            get
            {
                return stripFields;
            }
            set 
            {
                this.stripFields = value; 
            }
        }

        // this abstract class does not initialize the static field that are required for the the StripFields method.
        internal static Field SHEET_NAME_FIELD { get { return Fields.Instance.SHEET_NAME_FIELD; } }
        internal static Field DATE_FIELD { get { return Fields.Instance.DATE_FIELD; } }
        internal static Field FILE_FIELD { get { return Fields.Instance.FILE_FIELD; } }
        public static Field FULL_FILE_FIELD { get { return Fields.Instance.FULL_FILE_FIELD; } }
        internal static Field PAGE_FIELD { get { return Fields.Instance.PAGE_FIELD; } }
        internal static Field TIME_FIELD { get { return Fields.Instance.TIME_FIELD; } }
        internal static Field NUM_PAGES_FIELD { get { return Fields.Instance.NUM_PAGES_FIELD; } }

        public static Field PICTURE_FIELD { get { return Fields.Instance.PICTURE_FIELD; } }

        internal static PairField BOLD_FIELD { get { return Fields.Instance.BOLD_FIELD; } }
        public static PairField ITALIC_FIELD { get { return Fields.Instance.ITALIC_FIELD; } }
        public static PairField STRIKETHROUGH_FIELD { get { return Fields.Instance.STRIKETHROUGH_FIELD; } }
        public static PairField SUBSCRIPT_FIELD { get { return Fields.Instance.SUBSCRIPT_FIELD; } }
        public static PairField SUPERSCRIPT_FIELD { get { return Fields.Instance.SUPERSCRIPT_FIELD; } }
        internal static PairField UNDERLINE_FIELD { get { return Fields.Instance.UNDERLINE_FIELD; } }
        internal static PairField DOUBLE_UNDERLINE_FIELD { get { return Fields.Instance.DOUBLE_UNDERLINE_FIELD; } }


        /// <summary>
        /// Represents a special field in a header or footer,
        /// eg the page number
        /// </summary>
        public class Field
        {
            [Obsolete("Use the generic list Fields.AllFields instead.")]
            public static ArrayList ALL_FIELDS { get { return new ArrayList(Fields.AllFields); } }

            /** The character sequence that marks this field */
            public String sequence;
            public Field(Fields fields, String sequence)
            {
                this.sequence = sequence;
                fields.Add(this);
            }


        }
        /// <summary>
        /// A special field that normally comes in a pair, eg
        /// turn on underline / turn off underline
        /// </summary>
        public class PairField : Field
        {
            public PairField(Fields fields, String sequence)
                : base(fields, sequence)
            {

            }
        }


        public class Fields
        {
            private List<Field> allFields = new List<Field>();
            public static ReadOnlyCollection<Field> AllFields { get { return Instance.allFields.AsReadOnly(); } }
            private Field _sheetnamefield;
            private Field _filefield;
            private Field _fullfilefield;
            private Field _pagefield;
            private Field _datefield;
            private Field _timefield;
            private Field _numpagesfield;
            private Field _picturefield;
            private PairField _boldfield;
            private PairField _italicfield;
            private PairField _strikethroughfield;
            private PairField _subscriptfield;
            private PairField _superscriptfield;
            private PairField _underlinefield;
            private PairField _doubleunderlinefield;
            public Field SHEET_NAME_FIELD
            {
                get{return _sheetnamefield;}
            }
            public Field DATE_FIELD
            {
                get{return _datefield;}
            }

            public Field FILE_FIELD
            {
                get{return _filefield;}
            }
            public Field FULL_FILE_FIELD
            {
                get{return _fullfilefield;}
            }

            public Field PAGE_FIELD
            {
                get{return _pagefield;}
            }
            public Field TIME_FIELD
            {
                get{return _timefield;}
            }
            public Field NUM_PAGES_FIELD
            {
                get{return _numpagesfield;}
            }

            public Field PICTURE_FIELD
            {
                get{return _picturefield;}
            }

            public PairField BOLD_FIELD
            {
                get{return _boldfield;}
            }
            public PairField ITALIC_FIELD
            {
                get{return _italicfield;}
            }
            public PairField STRIKETHROUGH_FIELD
            {
                get{return _strikethroughfield;}
            }
            public PairField SUBSCRIPT_FIELD
            {
                get{return _subscriptfield;}
            }
            public PairField SUPERSCRIPT_FIELD
            {
                get{return _superscriptfield;}
            }
            public PairField UNDERLINE_FIELD
            {
                get{return _underlinefield;}
            }
            public PairField DOUBLE_UNDERLINE_FIELD
            {
                get{return _doubleunderlinefield;}
            }


            #region Singleton Implementation
            /// <summary>
            /// Instance to this class.
            /// </summary>
            static private readonly Fields instance = new Fields();

            /// <summary>
            /// Explicit static constructor to tell C# compiler not to mark type as beforefieldinit.
            /// </summary>
            static Fields()
            { }

            /// <summary>
            /// Initialize AllFields.
            /// </summary>
            private Fields()
            {
                _sheetnamefield = new Field(this, "&A");
                _datefield = new Field(this, "&D");
                _filefield  = new Field(this, "&F");
                _fullfilefield  = new Field(this, "&Z");
                _pagefield = new Field(this, "&P");
                _timefield = new Field(this, "&T");
                _numpagesfield = new Field(this, "&N");

                _picturefield = new Field(this, "&G");

                _boldfield = new PairField(this, "&B");
                _italicfield = new PairField(this, "&I");
                _strikethroughfield = new PairField(this, "&S");
                _subscriptfield = new PairField(this, "&Y");
                _superscriptfield = new PairField(this, "&X");
                _underlinefield = new PairField(this, "&U");
                _doubleunderlinefield = new PairField(this, "&E");
            }

            /// <summary>
            /// Accessing the initialized instance.
            /// </summary>
            public static Fields Instance
            {
                get
                {
                    return instance;
                }
            }
            #endregion Singleton Implementation

            internal void Add(Field field)
            {
                allFields.Add(field);
            }
        }


    }
}