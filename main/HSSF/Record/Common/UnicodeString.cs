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

namespace NPOI.HSSF.Record
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using NPOI.HSSF.Record.Cont;
    using NPOI.Util;

    /**
     * Title: Unicode String<p/>
     * Description:  Unicode String - just standard fields that are in several records.
     *               It is considered more desirable then repeating it in all of them.<p/>
     *               This is often called a XLUnicodeRichExtendedString in MS documentation.<p/>
     * REFERENCE:  PG 264 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)<p/>
     * REFERENCE:  PG 951 Excel Binary File Format (.xls) Structure Specification v20091214 
     */
    public class UnicodeString : IComparable<UnicodeString>
    { // TODO - make this when the compatibility version is Removed
        private static POILogger _logger = POILogFactory.GetLogger(typeof(UnicodeString));

        private short field_1_charCount;
        private byte field_2_optionflags;
        private String field_3_string;
        private List<FormatRun> field_4_format_Runs;
        private ExtRst field_5_ext_rst;
        private static BitField highByte = BitFieldFactory.GetInstance(0x1);
        // 0x2 is reserved
        private static BitField extBit = BitFieldFactory.GetInstance(0x4);
        private static BitField richText = BitFieldFactory.GetInstance(0x8);

        public class FormatRun : IComparable<FormatRun>
        {
            internal short _character;
            internal short _fontIndex;

            public FormatRun(short character, short fontIndex)
            {
                this._character = character;
                this._fontIndex = fontIndex;
            }

            public FormatRun(ILittleEndianInput in1) :
                this(in1.ReadShort(), in1.ReadShort())
            {
            }

            public short CharacterPos
            {
                get
                {
                    return _character;
                }
            }

            public short FontIndex
            {
                get
                {
                    return _fontIndex;
                }
            }

            public override bool Equals(Object o)
            {
                if (!(o is FormatRun))
                {
                    return false;
                }
                FormatRun other = (FormatRun)o;

                return _character == other._character && _fontIndex == other._fontIndex;
            }
            public override int GetHashCode()
            {
                return base.GetHashCode();
            }
            public int CompareTo(FormatRun r)
            {
                if (_character == r._character && _fontIndex == r._fontIndex)
                {
                    return 0;
                }
                if (_character == r._character)
                {
                    return _fontIndex - r._fontIndex;
                }
                return _character - r._character;
            }

            public override String ToString()
            {
                return "character=" + _character + ",fontIndex=" + _fontIndex;
            }

            public void Serialize(ILittleEndianOutput out1)
            {
                out1.WriteShort(_character);
                out1.WriteShort(_fontIndex);
            }
        }

        // See page 681
        public class ExtRst : IComparable<ExtRst>
        {
            private short reserved;

            // This is a Phs (see page 881)
            private short formattingFontIndex;
            private short formattingOptions;

            // This is a RPHSSub (see page 894)
            private int numberOfRuns;
            private String phoneticText;

            // This is an array of PhRuns (see page 881)
            private PhRun[] phRuns;
            // Sometimes there's some cruft at the end
            private byte[] extraData;

            private void populateEmpty()
            {
                reserved = 1;
                phoneticText = "";
                phRuns = new PhRun[0];
                extraData = new byte[0];
            }
            public override int GetHashCode()
            {
                int hash = reserved;
                hash = 31 * hash + formattingFontIndex;
                hash = 31 * hash + formattingOptions;
                hash = 31 * hash + numberOfRuns;
                hash = 31 * hash + phoneticText.GetHashCode();

                if (phRuns != null)
                {
                    foreach (PhRun ph in phRuns)
                    {
                        hash = 31 * hash + ph.phoneticTextFirstCharacterOffset;
                        hash = 31 * hash + ph.realTextFirstCharacterOffset;
                        hash = 31 * hash + ph.realTextLength;
                    }
                }
                return hash;
            }
            internal ExtRst()
            {
                populateEmpty();
            }
            internal ExtRst(ILittleEndianInput in1, int expectedLength)
            {
                reserved = in1.ReadShort();

                // Old style detection (Reserved = 0xFF)
                if (reserved == -1)
                {
                    populateEmpty();
                    return;
                }

                // Spot corrupt records
                if (reserved != 1)
                {
                    _logger.Log(POILogger.WARN, "Warning - ExtRst has wrong magic marker, expecting 1 but found " + reserved + " - ignoring");
                    // Grab all the remaining data, and ignore it
                    for (int i = 0; i < expectedLength - 2; i++)
                    {
                        in1.ReadByte();
                    }
                    // And make us be empty
                    populateEmpty();
                    return;
                }

                // Carry on Reading in as normal
                short stringDataSize = in1.ReadShort();

                formattingFontIndex = in1.ReadShort();
                formattingOptions = in1.ReadShort();

                // RPHSSub
                numberOfRuns = in1.ReadUShort();
                short length1 = in1.ReadShort();
                // No really. Someone Clearly forgot to read
                //  the docs on their datastructure...
                short length2 = in1.ReadShort();
                // And sometimes they write out garbage :(
                if (length1 == 0 && length2 > 0)
                {
                    length2 = 0;
                }
                if (length1 != length2)
                {
                    throw new InvalidOperationException(
                          "The two length fields of the Phonetic Text don't agree! " +
                          length1 + " vs " + length2
                    );
                }
                phoneticText = StringUtil.ReadUnicodeLE(in1, length1);

                int RunData = stringDataSize - 4 - 6 - (2 * phoneticText.Length);
                int numRuns = (RunData / 6);
                phRuns = new PhRun[numRuns];
                for (int i = 0; i < phRuns.Length; i++)
                {
                    phRuns[i] = new PhRun(in1);
                }

                int extraDataLength = RunData - (numRuns * 6);
                if (extraDataLength < 0)
                {
                    //System.err.Println("Warning - ExtRst overran by " + (0-extraDataLength) + " bytes");
                    extraDataLength = 0;
                }
                extraData = new byte[extraDataLength];
                for (int i = 0; i < extraData.Length; i++)
                {
                    extraData[i] = (byte)in1.ReadByte();
                }
            }
            /**
             * Returns our size, excluding our 
             *  4 byte header
             */
            internal int DataSize
            {
                get
                {
                    return 4 + 6 + (2 * phoneticText.Length) +
                       (6 * phRuns.Length) + extraData.Length;
                }
            }
            internal void Serialize(ContinuableRecordOutput out1)
            {
                int dataSize = DataSize;

                out1.WriteContinueIfRequired(8);
                out1.WriteShort(reserved);
                out1.WriteShort(dataSize);
                out1.WriteShort(formattingFontIndex);
                out1.WriteShort(formattingOptions);

                out1.WriteContinueIfRequired(6);
                out1.WriteShort(numberOfRuns);
                out1.WriteShort(phoneticText.Length);
                out1.WriteShort(phoneticText.Length);

                out1.WriteContinueIfRequired(phoneticText.Length * 2);
                StringUtil.PutUnicodeLE(phoneticText, out1);

                for (int i = 0; i < phRuns.Length; i++)
                {
                    phRuns[i].Serialize(out1);
                }

                out1.Write(extraData);
            }

            public override bool Equals(Object obj)
            {
                if (!(obj is ExtRst))
                {
                    return false;
                }
                ExtRst other = (ExtRst)obj;
                return (CompareTo(other) == 0);
            }
            public override string ToString()
            {
                return base.ToString();
            }
                 
            public int CompareTo(ExtRst o)
            {
                int result;

                result = reserved - o.reserved;
                if (result != 0) return result;
                result = formattingFontIndex - o.formattingFontIndex;
                if (result != 0) return result;
                result = formattingOptions - o.formattingOptions;
                if (result != 0) return result;
                result = numberOfRuns - o.numberOfRuns;
                if (result != 0) return result;

                //result = phoneticText.CompareTo(o.phoneticText);
                result = string.Compare(phoneticText, o.phoneticText, StringComparison.CurrentCulture);
                if (result != 0) return result;

                result = phRuns.Length - o.phRuns.Length;
                if (result != 0) return result;
                for (int i = 0; i < phRuns.Length; i++)
                {
                    result = phRuns[i].phoneticTextFirstCharacterOffset - o.phRuns[i].phoneticTextFirstCharacterOffset;
                    if (result != 0) return result;
                    result = phRuns[i].realTextFirstCharacterOffset - o.phRuns[i].realTextFirstCharacterOffset;
                    if (result != 0) return result;
                    result = phRuns[i].realTextLength - o.phRuns[i].realTextLength;
                    if (result != 0) return result;
                }

                result = Arrays.HashCode(extraData) - Arrays.HashCode(o.extraData);

                // If we Get here, it's the same
                return result;
            }

            internal ExtRst Clone()
            {
                ExtRst ext = new ExtRst();
                ext.reserved = reserved;
                ext.formattingFontIndex = formattingFontIndex;
                ext.formattingOptions = formattingOptions;
                ext.numberOfRuns = numberOfRuns;
                ext.phoneticText = phoneticText;
                ext.phRuns = new PhRun[phRuns.Length];
                for (int i = 0; i < ext.phRuns.Length; i++)
                {
                    ext.phRuns[i] = new PhRun(
                          phRuns[i].phoneticTextFirstCharacterOffset,
                          phRuns[i].realTextFirstCharacterOffset,
                          phRuns[i].realTextLength
                    );
                }
                return ext;
            }

            public short FormattingFontIndex
            {
                get
                {
                    return formattingFontIndex;
                }
            }
            public short FormattingOptions
            {
                get
                {
                    return formattingOptions;
                }
            }
            public int NumberOfRuns
            {
                get
                {
                    return numberOfRuns;
                }
            }
            public String PhoneticText
            {
                get
                {
                    return phoneticText;
                }
            }
            public PhRun[] PhRuns
            {
                get
                {
                    return phRuns;
                }
            }
        }
        public class PhRun
        {
            internal int phoneticTextFirstCharacterOffset;
            internal int realTextFirstCharacterOffset;
            internal int realTextLength;

            public PhRun(int phoneticTextFirstCharacterOffset,
                 int realTextFirstCharacterOffset, int realTextLength)
            {
                this.phoneticTextFirstCharacterOffset = phoneticTextFirstCharacterOffset;
                this.realTextFirstCharacterOffset = realTextFirstCharacterOffset;
                this.realTextLength = realTextLength;
            }
            internal PhRun(ILittleEndianInput in1)
            {
                phoneticTextFirstCharacterOffset = in1.ReadUShort();
                realTextFirstCharacterOffset = in1.ReadUShort();
                realTextLength = in1.ReadUShort();
            }
            internal void Serialize(ContinuableRecordOutput out1)
            {
                out1.WriteContinueIfRequired(6);
                out1.WriteShort(phoneticTextFirstCharacterOffset);
                out1.WriteShort(realTextFirstCharacterOffset);
                out1.WriteShort(realTextLength);
            }
        }

        private UnicodeString()
        {
            //Used for clone method.
        }

        public UnicodeString(String str)
        {
            String = (str);
        }



        public override int GetHashCode()
        {
            int stringHash = 0;
            if (field_3_string != null)
                stringHash = field_3_string.GetHashCode();
            return field_1_charCount + stringHash;
        }

        /**
         * Our handling of Equals is inconsistent with CompareTo.  The trouble is because we don't truely understand
         * rich text fields yet it's difficult to make a sound comparison.
         *
         * @param o     The object to Compare.
         * @return      true if the object is actually Equal.
         */
        public override bool Equals(Object o)
        {
            if (!(o is UnicodeString))
            {
                return false;
            }
            UnicodeString other = (UnicodeString)o;

            //OK lets do this in stages to return a quickly, first check the actual string
            bool eq = ((field_1_charCount == other.field_1_charCount)
                    && (field_2_optionflags == other.field_2_optionflags)
                    && field_3_string.Equals(other.field_3_string));
            if (!eq) return false;

            //OK string appears to be equal but now lets compare formatting Runs
            if ((field_4_format_Runs == null) && (other.field_4_format_Runs == null))
                //Strings are Equal, and there are not formatting Runs.
                return true;
            if (((field_4_format_Runs == null) && (other.field_4_format_Runs != null)) ||
                 (field_4_format_Runs != null) && (other.field_4_format_Runs == null))
                //Strings are Equal, but one or the other has formatting Runs
                return false;

            //Strings are Equal, so now compare formatting Runs.
            int size = field_4_format_Runs.Count;
            if (size != other.field_4_format_Runs.Count)
                return false;

            for (int i = 0; i < size; i++)
            {
                FormatRun Run1 = field_4_format_Runs[(i)];
                FormatRun run2 = other.field_4_format_Runs[(i)];

                if (!Run1.Equals(run2))
                    return false;
            }

            // Well the format Runs are equal as well!, better check the ExtRst data
            if (field_5_ext_rst == null && other.field_5_ext_rst == null)
            {
                // Good
            }
            else if (field_5_ext_rst != null && other.field_5_ext_rst != null)
            {
                int extCmp = field_5_ext_rst.CompareTo(other.field_5_ext_rst);
                if (extCmp == 0)
                {
                    // Good
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

            //Phew!! After all of that we have finally worked out that the strings
            //are identical.
            return true;
        }

        /**
         * construct a unicode string record and fill its fields, ID is ignored
         * @param in the RecordInputstream to read the record from
         */
        public UnicodeString(RecordInputStream in1)
        {
            field_1_charCount = in1.ReadShort();
            field_2_optionflags = (byte)in1.ReadByte();

            int RunCount = 0;
            int extensionLength = 0;
            //Read the number of rich Runs if rich text.
            if (IsRichText)
            {
                RunCount = in1.ReadShort();
            }
            //Read the size of extended data if present.
            if (IsExtendedText)
            {
                extensionLength = in1.ReadInt();
            }

            bool IsCompressed = ((field_2_optionflags & 1) == 0);
            if (IsCompressed)
            {
                field_3_string = in1.ReadCompressedUnicode(CharCount);
            }
            else
            {
                field_3_string = in1.ReadUnicodeLEString(CharCount);
            }


            if (IsRichText && (RunCount > 0))
            {
                field_4_format_Runs = new List<FormatRun>(RunCount);
                for (int i = 0; i < RunCount; i++)
                {
                    field_4_format_Runs.Add(new FormatRun(in1));
                }
            }

            if (IsExtendedText && (extensionLength > 0))
            {
                field_5_ext_rst = new ExtRst(new ContinuableRecordInput(in1), extensionLength);
                if (field_5_ext_rst.DataSize + 4 != extensionLength)
                {
                    _logger.Log(POILogger.WARN, "ExtRst was supposed to be " + extensionLength + " bytes long, but seems to actually be " + (field_5_ext_rst.DataSize + 4));
                }
            }
        }



        /**
             * get the number of characters in the string,
             *  as an un-wrapped int
             *
             * @return number of characters
             */
        public int CharCount
        {
            get
            {
                if (field_1_charCount < 0)
                {
                    return field_1_charCount + 65536;
                }
                return field_1_charCount;
            }
            set
            {
                field_1_charCount = (short)value;
            }
        }
        public short CharCountShort
        {
            get { return field_1_charCount; }
        }

        /**
         * Get the option flags which among other things return if this is a 16-bit or
         * 8 bit string
         *
         * @return optionflags bitmask
         *
         */

        public byte OptionFlags
        {
            get
            {
                return field_2_optionflags;
            }
            set
            {
                field_2_optionflags = value;
            }
        }


        /**
             * @return the actual string this Contains as a java String object
             */
        public String String
        {
            get
            {
                return field_3_string;
            }
            set
            {
                field_3_string = value;
                CharCount = ((short)field_3_string.Length);
                // scan for characters greater than 255 ... if any are
                // present, we have to use 16-bit encoding. Otherwise, we
                // can use 8-bit encoding
                bool useUTF16 = false;
                int strlen = value.Length;

                for (int j = 0; j < strlen; j++)
                {
                    if (value[j] > 255)
                    {
                        useUTF16 = true;
                        break;
                    }
                }
                if (useUTF16)
                    //Set the uncompressed bit
                    field_2_optionflags = highByte.SetByte(field_2_optionflags);
                else field_2_optionflags = highByte.ClearByte(field_2_optionflags);

            }
        }

        public int FormatRunCount
        {
            get
            {
                if (field_4_format_Runs == null)
                    return 0;
                return field_4_format_Runs.Count;
            }
        }

        public FormatRun GetFormatRun(int index)
        {
            if (field_4_format_Runs == null)
            {
                return null;
            }
            if (index < 0 || index >= field_4_format_Runs.Count)
            {
                return null;
            }
            return field_4_format_Runs[(index)];
        }

        private int FindFormatRunAt(int characterPos)
        {
            int size = field_4_format_Runs.Count;
            for (int i = 0; i < size; i++)
            {
                FormatRun r = field_4_format_Runs[(i)];
                if (r._character == characterPos)
                    return i;
                else if (r._character > characterPos)
                    return -1;
            }
            return -1;
        }

        /** Adds a font run to the formatted string.
         *
         *  If a font run exists at the current charcter location, then it is
         *  Replaced with the font run to be Added.
         */
        public void AddFormatRun(FormatRun r)
        {
            if (field_4_format_Runs == null)
            {
                field_4_format_Runs = new List<FormatRun>();
            }

            int index = FindFormatRunAt(r._character);
            if (index != -1)
                field_4_format_Runs.RemoveAt(index);

            field_4_format_Runs.Add(r);
            //Need to sort the font Runs to ensure that the font Runs appear in
            //character order
            //collections.Sort(field_4_format_Runs);
            field_4_format_Runs.Sort();

            //Make sure that we now say that we are a rich string
            field_2_optionflags = richText.SetByte(field_2_optionflags);
        }

        public List<FormatRun> FormatIterator()
        {
            if (field_4_format_Runs != null)
            {
                return field_4_format_Runs;
            }
            return null;
        }

        public void RemoveFormatRun(FormatRun r)
        {
            field_4_format_Runs.Remove(r);
            if (field_4_format_Runs.Count == 0)
            {
                field_4_format_Runs = null;
                field_2_optionflags = richText.ClearByte(field_2_optionflags);
            }
        }

        public void ClearFormatting()
        {
            field_4_format_Runs = null;
            field_2_optionflags = richText.ClearByte(field_2_optionflags);
        }


        public ExtRst ExtendedRst
        {
            get
            {
                return this.field_5_ext_rst;
            }
            set
            {
                if (value != null)
                {
                    field_2_optionflags = extBit.SetByte(field_2_optionflags);
                }
                else
                {
                    field_2_optionflags = extBit.ClearByte(field_2_optionflags);
                }
                this.field_5_ext_rst = value;
            }
        }



        /**
         * Swaps all use in the string of one font index
         *  for use of a different font index.
         * Normally only called when fonts have been
         *  Removed / re-ordered
         */
        public void SwapFontUse(short oldFontIndex, short newFontIndex)
        {
            foreach (FormatRun run in field_4_format_Runs)
            {
                if (run._fontIndex == oldFontIndex)
                {
                    run._fontIndex = newFontIndex;
                }
            }
        }

        /**
         * unlike the real records we return the same as "getString()" rather than debug info
         * @see #getDebugInfo()
         * @return String value of the record
         */

        public override String ToString()
        {
            return String;
        }

        /**
         * return a character representation of the fields of this record
         *
         *
         * @return String of output for biffviewer etc.
         *
         */
        public String GetDebugInfo()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[UNICODESTRING]\n");
            buffer.Append("    .charcount       = ")
                .Append(StringUtil.ToHexString(CharCount)).Append("\n");
            buffer.Append("    .optionflags     = ")
                .Append(StringUtil.ToHexString(OptionFlags)).Append("\n");
            buffer.Append("    .string          = ").Append(String).Append("\n");
            if (field_4_format_Runs != null)
            {
                for (int i = 0; i < field_4_format_Runs.Count; i++)
                {
                    FormatRun r = field_4_format_Runs[(i)];
                    buffer.Append("      .format_Run" + i + "          = ").Append(r.ToString()).Append("\n");
                }
            }
            if (field_5_ext_rst != null)
            {
                buffer.Append("    .field_5_ext_rst          = ").Append("\n");
                buffer.Append(field_5_ext_rst.ToString()).Append("\n");
            }
            buffer.Append("[/UNICODESTRING]\n");
            return buffer.ToString();
        }

        /**
         * Serialises out the String. There are special rules
         *  about where we can and can't split onto
         *  Continue records.
         */
        public void Serialize(ContinuableRecordOutput out1)
        {
            int numberOfRichTextRuns = 0;
            int extendedDataSize = 0;
            if (IsRichText && field_4_format_Runs != null)
            {
                numberOfRichTextRuns = field_4_format_Runs.Count;
            }
            if (IsExtendedText && field_5_ext_rst != null)
            {
                extendedDataSize = 4 + field_5_ext_rst.DataSize;
            }

            // Serialise the bulk of the String
            // The WriteString handles tricky continue stuff for us
            out1.WriteString(field_3_string, numberOfRichTextRuns, extendedDataSize);

            if (numberOfRichTextRuns > 0)
            {

                //This will ensure that a run does not split a continue
                for (int i = 0; i < numberOfRichTextRuns; i++)
                {
                    if (out1.AvailableSpace < 4)
                    {
                        out1.WriteContinue();
                    }
                    FormatRun r = field_4_format_Runs[(i)];
                    r.Serialize(out1);
                }
            }

            if (extendedDataSize > 0)
            {
                field_5_ext_rst.Serialize(out1);
            }
        }

        public int CompareTo(UnicodeString str)
        {

            //int result = String.CompareTo(str.String);
            int result = string.Compare(String, str.String, StringComparison.CurrentCulture);

            //As per the Equals method lets do this in stages
            if (result != 0)
                return result;

            //OK string appears to be equal but now lets compare formatting Runs
            if ((field_4_format_Runs == null) && (str.field_4_format_Runs == null))
                //Strings are Equal, and there are no formatting Runs.
                return 0;

            if ((field_4_format_Runs == null) && (str.field_4_format_Runs != null))
                //Strings are Equal, but one or the other has formatting Runs
                return 1;
            if ((field_4_format_Runs != null) && (str.field_4_format_Runs == null))
                //Strings are Equal, but one or the other has formatting Runs
                return -1;

            //Strings are Equal, so now compare formatting Runs.
            int size = field_4_format_Runs.Count;
            if (size != str.field_4_format_Runs.Count)
                return size - str.field_4_format_Runs.Count;

            for (int i = 0; i < size; i++)
            {
                FormatRun Run1 = field_4_format_Runs[(i)];
                FormatRun run2 = str.field_4_format_Runs[(i)];

                result = Run1.CompareTo(run2);
                if (result != 0)
                    return result;
            }

            //Well the format Runs are equal as well!, better check the ExtRst data
            if ((field_5_ext_rst == null) && (str.field_5_ext_rst == null))
                return 0;
            if ((field_5_ext_rst == null) && (str.field_5_ext_rst != null))
                return 1;
            if ((field_5_ext_rst != null) && (str.field_5_ext_rst == null))
                return -1;

            result = field_5_ext_rst.CompareTo(str.field_5_ext_rst);
            if (result != 0)
                return result;

            //Phew!! After all of that we have finally worked out that the strings
            //are identical.
            return 0;
        }

        private bool IsRichText
        {
            get
            {
                return richText.IsSet(OptionFlags);
            }
        }

        private bool IsExtendedText
        {
            get
            {
                return extBit.IsSet(OptionFlags);
            }
        }

        public Object Clone()
        {
            UnicodeString str = new UnicodeString();
            str.field_1_charCount = field_1_charCount;
            str.field_2_optionflags = field_2_optionflags;
            str.field_3_string = field_3_string;
            if (field_4_format_Runs != null)
            {
                str.field_4_format_Runs = new List<FormatRun>();
                foreach (FormatRun r in field_4_format_Runs)
                {
                    str.field_4_format_Runs.Add(new FormatRun(r._character, r._fontIndex));
                }
            }
            if (field_5_ext_rst != null)
            {
                str.field_5_ext_rst = field_5_ext_rst.Clone();
            }

            return str;
        }
    }

}
