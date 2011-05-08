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
    using System.Text;
    using NPOI.HSSF.Record.Cont;
    using NPOI.Util;
    using NPOI.Util.IO;
    using System.Collections.Generic;
    public class FormatRun : IComparable<FormatRun>
    {
        internal short _character;
        internal short _fontIndex;

        public FormatRun(short character, short fontIndex)
        {
            this._character = character;
            this._fontIndex = fontIndex;
        }

        public FormatRun(LittleEndianInput in1)
            : this(in1.ReadShort(), in1.ReadShort())
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

        public void Serialize(LittleEndianOutput out1)
        {
            out1.WriteShort(_character);
            out1.WriteShort(_fontIndex);
        }
    }
    /**
     * Title: Unicode String<p/>
     * Description:  Unicode String - just standard fields that are in several records.
     *               It is considered more desirable then repeating it in all of them.<p/>
     * REFERENCE:  PG 264 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)<p/>
     * @author  Andrew C. Oliver
     * @author Marc Johnson (mjohnson at apache dot org)
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class UnicodeString : IComparable
    {
        private short field_1_charCount;
        private byte field_2_optionflags;
        private String field_3_string;
        private List<FormatRun> field_4_format_Runs;
        private byte[] field_5_ext_rst;
        private static BitField highByte = BitFieldFactory.GetInstance(0x1);
        private static BitField extBit = BitFieldFactory.GetInstance(0x4);
        private static BitField richText = BitFieldFactory.GetInstance(0x8);



        private UnicodeString()
        {
            //Used for clone method.
        }

        public UnicodeString(String str)
        {
            this.String = str;
        }



        public override int GetHashCode()
        {
            int stringHash = 0;
            if (field_3_string != null)
                stringHash = field_3_string.GetHashCode();
            return field_1_charCount + stringHash;
        }

        /**
         * Our handling of equals is inconsistent with CompareTo.  The trouble is because we don't truely understand
         * rich text fields yet it's difficult to make a sound comparison.
         *
         * @param o     The object to Compare.
         * @return      true if the object is actually equal.
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
                //Strings are equal, and there are not formatting Runs.
                return true;
            if (((field_4_format_Runs == null) && (other.field_4_format_Runs != null)) ||
                 (field_4_format_Runs != null) && (other.field_4_format_Runs == null))
                //Strings are equal, but one or the other has formatting Runs
                return false;

            //Strings are equal, so now compare formatting Runs.
            int size = field_4_format_Runs.Count;
            if (size != other.field_4_format_Runs.Count)
                return false;

            for (int i = 0; i < size; i++)
            {
                FormatRun Run1 = field_4_format_Runs[i];
                FormatRun Run2 = other.field_4_format_Runs[i];

                if (!Run1.Equals(Run2))
                    return false;
            }

            //Well the format Runs are equal as well!, better check the ExtRst data
            //Which by the way we dont know how to decode!
            if ((field_5_ext_rst == null) && (other.field_5_ext_rst == null))
                return true;
            if (((field_5_ext_rst == null) && (other.field_5_ext_rst != null)) ||
                ((field_5_ext_rst != null) && (other.field_5_ext_rst == null)))
                return false;
            size = field_5_ext_rst.Length;
            if (size != field_5_ext_rst.Length)
                return false;

            //Check individual bytes!
            for (int i = 0; i < size; i++)
            {
                if (field_5_ext_rst[i] != other.field_5_ext_rst[i])
                    return false;
            }
            //Phew!! After all of that we have finally worked out that the strings
            //are identical.
            return true;
        }

        /**
         * construct a unicode string record and fill its fields, ID is ignored
         * @param in the RecordInPutstream to read the record from
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

            bool isCompressed = ((field_2_optionflags & 1) == 0);
            if (isCompressed)
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
                field_5_ext_rst = new byte[extensionLength];
                for (int i = 0; i < extensionLength; i++)
                {
                    field_5_ext_rst[i] = (byte)in1.ReadByte();
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
         * get the option flags which among other things return if this is a 16-bit or
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
            return field_4_format_Runs[index];
        }

        private int FindFormatRunAt(int characterPos)
        {
            int size = field_4_format_Runs.Count;
            for (int i = 0; i < size; i++)
            {
                FormatRun r = field_4_format_Runs[i];
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


        public void SetExtendedRst(byte[] ext_rst)
        {
            if (ext_rst != null)
                field_2_optionflags = extBit.SetByte(field_2_optionflags);
            else field_2_optionflags = extBit.ClearByte(field_2_optionflags);
            this.field_5_ext_rst = ext_rst;
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
         * unlike the real records we return the same as "GetString()" rather than debug info
         * @see #GetDebugInfo()
         * @return String value of the record
         */

        public override String ToString()
        {
            return this.String;
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
                    FormatRun r = field_4_format_Runs[i];
                    buffer.Append("      .format_Run" + i + "          = ").Append(r.ToString()).Append("\n");
                }
            }
            if (field_5_ext_rst != null)
            {
                buffer.Append("    .field_5_ext_rst          = ").Append("\n").Append(HexDump.ToHex(field_5_ext_rst)).Append("\n");
            }
            buffer.Append("[/UNICODESTRING]\n");
            return buffer.ToString();
        }

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
                extendedDataSize = field_5_ext_rst.Length;
            }

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
                    FormatRun r = field_4_format_Runs[i];
                    r.Serialize(out1);
                }
            }

            if (extendedDataSize > 0)
            {
                // OK ExtRst is actually not documented, so i am going to hope
                // that we can actually continue on byte boundaries

                int extPos = 0;
                while (true)
                {
                    int nBytesToWrite = Math.Min(extendedDataSize - extPos, out1.AvailableSpace);
                    out1.Write(field_5_ext_rst, extPos, nBytesToWrite);
                    extPos += nBytesToWrite;
                    if (extPos >= extendedDataSize)
                    {
                        break;
                    }
                    out1.WriteContinue();
                }
            }
        }

        public int CompareTo(object obj)
        {
            UnicodeString str = (UnicodeString)obj;
            int result = this.String.CompareTo(str.String);

            //As per the equals method lets do this in stages
            if (result != 0)
                return result;

            //OK string appears to be equal but now lets compare formatting Runs
            if ((field_4_format_Runs == null) && (str.field_4_format_Runs == null))
                //Strings are equal, and there are no formatting Runs.
                return 0;

            if ((field_4_format_Runs == null) && (str.field_4_format_Runs != null))
                //Strings are equal, but one or the other has formatting Runs
                return 1;
            if ((field_4_format_Runs != null) && (str.field_4_format_Runs == null))
                //Strings are equal, but one or the other has formatting Runs
                return -1;

            //Strings are equal, so now compare formatting Runs.
            int size = field_4_format_Runs.Count;
            if (size != str.field_4_format_Runs.Count)
                return size - str.field_4_format_Runs.Count;

            for (int i = 0; i < size; i++)
            {
                FormatRun Run1 = field_4_format_Runs[i];
                FormatRun Run2 = str.field_4_format_Runs[i];

                result = Run1.CompareTo(Run2);
                if (result != 0)
                    return result;
            }

            //Well the format Runs are equal as well!, better check the ExtRst data
            //Which by the way we don't know how to decode!
            if ((field_5_ext_rst == null) && (str.field_5_ext_rst == null))
                return 0;
            if ((field_5_ext_rst == null) && (str.field_5_ext_rst != null))
                return 1;
            if ((field_5_ext_rst != null) && (str.field_5_ext_rst == null))
                return -1;

            size = field_5_ext_rst.Length;
            if (size != field_5_ext_rst.Length)
                return size - field_5_ext_rst.Length;

            //Check individual bytes!
            for (int i = 0; i < size; i++)
            {
                if (field_5_ext_rst[i] != str.field_5_ext_rst[i])
                    return field_5_ext_rst[i] - str.field_5_ext_rst[i];
            }
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
                str.field_5_ext_rst = new byte[field_5_ext_rst.Length];
                Array.Copy(field_5_ext_rst, 0, str.field_5_ext_rst, 0,
                                 field_5_ext_rst.Length);
            }

            return str;
        }
    }
}