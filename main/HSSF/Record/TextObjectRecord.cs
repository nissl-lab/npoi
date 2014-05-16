/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License Is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using NPOI.HSSF.UserModel;

    using NPOI.SS.UserModel;
    using NPOI.HSSF.Record.Cont;
    using NPOI.SS.Formula.PTG;
    using System.Globalization;

    public class TextObjectRecord : ContinuableRecord
    {
        NPOI.SS.UserModel.IRichTextString _text;

        public const short sid = 0x1B6;

        private const int FORMAT_RUN_ENCODED_SIZE = 8; // 2 shorts and 4 bytes reserved

        private BitField _HorizontalTextAlignment = BitFieldFactory.GetInstance(0x000E);
        private BitField _VerticalTextAlignment = BitFieldFactory.GetInstance(0x0070);
        private BitField textLocked = BitFieldFactory.GetInstance(0x200);

        private int field_1_options;
        private int field_2_textOrientation;
        private int field_3_reserved4;
        private int field_4_reserved5;
        private int field_5_reserved6;
        private int field_8_reserved7;
        /*
         * Note - the next three fields are very similar to those on
         * EmbededObjectRefSubRecord(ftPictFmla 0x0009)
         * 
         * some observed values for the 4 bytes preceding the formula: C0 5E 86 03
         * C0 11 AC 02 80 F1 8A 03 D4 F0 8A 03
         */
        private int _unknownPreFormulaInt;
        /** expect tRef, tRef3D, tArea, tArea3D or tName */
        private OperandPtg _linkRefPtg;
        /**
         * Not clear if needed .  Excel seems to be OK if this byte is not present. 
         * Value is often the same as the earlier firstColumn byte. */
        private Byte? _unknownPostFormulaByte;

        public TextObjectRecord()
        {

        }

        public TextObjectRecord(RecordInputStream in1)
        {

            field_1_options = in1.ReadUShort();
            field_2_textOrientation = in1.ReadUShort();
            field_3_reserved4 = in1.ReadUShort();
            field_4_reserved5 = in1.ReadUShort();
            field_5_reserved6 = in1.ReadUShort();
            int field_6_textLength = in1.ReadUShort();
            int field_7_formattingDataLength = in1.ReadUShort();
            field_8_reserved7 = in1.ReadInt();

            if (in1.Remaining > 0)
            {
                // Text Objects can have simple reference formulas
                // (This bit not mentioned in the MS document)
                if (in1.Remaining < 11)
                {
                    throw new RecordFormatException("Not enough remaining data for a link formula");
                }
                int formulaSize = in1.ReadUShort();
                _unknownPreFormulaInt = in1.ReadInt();
                Ptg[] ptgs = Ptg.ReadTokens(formulaSize, in1);
                if (ptgs.Length != 1)
                {
                    throw new RecordFormatException("Read " + ptgs.Length
                            + " tokens but expected exactly 1");
                }
                _linkRefPtg = (OperandPtg)ptgs[0];
                if (in1.Remaining > 0)
                {
                    _unknownPostFormulaByte = (byte)in1.ReadByte();
                }
                else
                {
                    _unknownPostFormulaByte = null;
                }
            }
            else
            {
                _linkRefPtg = null;
            }
            if (in1.Remaining > 0)
            {
                throw new RecordFormatException("Unused " + in1.Remaining + " bytes at end of record");
            }

            String text;
            if (field_6_textLength > 0)
            {
                text = ReadRawString(in1, field_6_textLength);
            }
            else
            {
                text = "";
            }
            _text = new HSSFRichTextString(text);

            if (field_7_formattingDataLength > 0)
            {
                ProcessFontRuns(in1, _text, field_7_formattingDataLength);
            }
        }
        private static void ProcessFontRuns(RecordInputStream in1, IRichTextString str,
            int formattingRunDataLength)
        {
            if (formattingRunDataLength % FORMAT_RUN_ENCODED_SIZE != 0)
            {
                throw new RecordFormatException("Bad format run data length " + formattingRunDataLength
                        + ")");
            }
            int nRuns = formattingRunDataLength / FORMAT_RUN_ENCODED_SIZE;
            for (int i = 0; i < nRuns; i++)
            {
                short index = in1.ReadShort();
                short iFont = in1.ReadShort();
                in1.ReadInt(); // skip reserved.
                str.ApplyFont(index, str.Length, iFont);
            }
        }

        private int TrailingRecordsSize
        {
            get
            {
                if (_text.Length < 1)
                {
                    return 0;
                }
                int encodedTextSize = 0;
                int textBytesLength = _text.Length * LittleEndianConsts.SHORT_SIZE;
                while (textBytesLength > 0)
                {
                    int chunkSize = Math.Min(RecordInputStream.MAX_RECORD_DATA_SIZE - 2, textBytesLength);
                    textBytesLength -= chunkSize;

                    encodedTextSize += 4;           // +4 for ContinueRecord sid+size
                    encodedTextSize += 1 + chunkSize; // +1 for compressed unicode flag, 
                }

                int encodedFormatSize = (_text.NumFormattingRuns + 1) * FORMAT_RUN_ENCODED_SIZE
                    + 4;  // +4 for ContinueRecord sid+size
                return encodedTextSize + encodedFormatSize;
            }
        }
        private static byte[] CreateFormatData(IRichTextString str)
        {
            int nRuns = str.NumFormattingRuns;
            byte[] result = new byte[(nRuns + 1) * FORMAT_RUN_ENCODED_SIZE];
            int pos = 0;
            for (int i = 0; i < nRuns; i++)
            {
                LittleEndian.PutUShort(result, pos, str.GetIndexOfFormattingRun(i));
                pos += 2;
                int fontIndex = ((HSSFRichTextString)str).GetFontOfFormattingRun(i);
                LittleEndian.PutUShort(result, pos, fontIndex == HSSFRichTextString.NO_FONT ? 0 : fontIndex);
                pos += 2;
                pos += 4; // skip reserved
            }
            LittleEndian.PutUShort(result, pos, str.Length);
            pos += 2;
            LittleEndian.PutUShort(result, pos, 0);
            pos += 2;
            pos += 4; // skip reserved

            return result;
        }

        private void SerializeTrailingRecords(ContinuableRecordOutput out1)
        {
            out1.WriteContinue();
            out1.WriteStringData(_text.String);
            out1.WriteContinue();
            WriteFormatData(out1,_text);
        }

        private void WriteFormatData(ContinuableRecordOutput out1, IRichTextString str)
        {
            int nRuns = str.NumFormattingRuns;
            for (int i = 0; i < nRuns; i++)
            {
                out1.WriteShort(str.GetIndexOfFormattingRun(i));
                int fontIndex = ((HSSFRichTextString)str).GetFontOfFormattingRun(i);
                out1.WriteShort(fontIndex == HSSFRichTextString.NO_FONT ? 0 : fontIndex);
                out1.WriteInt(0); // skip reserved
            }
            out1.WriteShort(str.Length);
            out1.WriteShort(0);
            out1.WriteInt(0); // skip reserved
        }


        private int FormattingDataLength
        {
            get
            {
                if (_text.Length < 1)
                {
                    // important - no formatting data if text is empty 
                    return 0;
                }
                return (_text.NumFormattingRuns + 1) * FORMAT_RUN_ENCODED_SIZE;
            }
        }

        private void SerializeTXORecord(ContinuableRecordOutput out1)
        {
            out1.WriteShort(field_1_options);
            out1.WriteShort(field_2_textOrientation);
            out1.WriteShort(field_3_reserved4);
            out1.WriteShort(field_4_reserved5);
            out1.WriteShort(field_5_reserved6);
            out1.WriteShort(_text.Length);
            out1.WriteShort(FormattingDataLength);
            out1.WriteInt(field_8_reserved7);

            if (_linkRefPtg != null)
            {
                int formulaSize = _linkRefPtg.Size;
                out1.WriteShort(formulaSize);
                out1.WriteInt(_unknownPreFormulaInt);
                _linkRefPtg.Write(out1);

                if (_unknownPostFormulaByte != null)
                {
                    out1.WriteByte(Convert.ToByte(_unknownPostFormulaByte, CultureInfo.InvariantCulture));
                }
            }
        }


        protected override void Serialize(ContinuableRecordOutput out1)
        {
            SerializeTXORecord(out1);

            if (_text.String.Length > 0)
            {
                SerializeTrailingRecords(out1);
            }
        }

        private void ProcessFontRuns(RecordInputStream in1)
        {
            while (in1.Remaining > 0)
            {
                short index = in1.ReadShort();
                short iFont = in1.ReadShort();
                in1.ReadInt();  // skip reserved.

                _text.ApplyFont(index, _text.Length, iFont);
            }
        }


        private static String ReadRawString(RecordInputStream in1, int textLength)
        {
            byte compressByte = (byte)in1.ReadByte();
            bool isCompressed = (compressByte & 0x01) == 0;
            if (isCompressed)
            {
                return in1.ReadCompressedUnicode(textLength);
            }
            return in1.ReadUnicodeLEString(textLength);
        }

        public IRichTextString Str
        {
            get { return _text; }
            set { this._text = value; }
        }
        public override short Sid
        {
            get
            {
                return sid;
            }
        }

        /**
         * Get the text orientation field for the TextObjectBase record.
         *
         * @return a TextOrientation
         */
        public TextOrientation TextOrientation
        {
            get { return (TextOrientation) field_2_textOrientation; }
            set { this.field_2_textOrientation = (int) value; }
        }


        /**
 * @return the Horizontal text alignment field value.
 */
        public HorizontalTextAlignment HorizontalTextAlignment
        {
            get {
                return (HorizontalTextAlignment)_HorizontalTextAlignment.GetValue(field_1_options);
            }
            set {
                field_1_options = _HorizontalTextAlignment.SetValue(field_1_options, (int) value);
            }
        }
        /**
 * @return the Vertical text alignment field value.
 */
        public VerticalTextAlignment VerticalTextAlignment
        {
            get {
                return (VerticalTextAlignment)_VerticalTextAlignment.GetValue(field_1_options);
            }
            set {
                field_1_options = _VerticalTextAlignment.SetValue(field_1_options, (int) value);
            }
        }

        /**
         * Text has been locked
         * @return  the text locked field value.
         */
        public bool IsTextLocked
        {
            get { return textLocked.IsSet(field_1_options); }
            set { field_1_options = textLocked.SetBoolean(field_1_options, value); }
        }

        public Ptg LinkRefPtg
        {
            get
            {
                return _linkRefPtg;
            }
        }
        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("[TXO]\n");
            sb.Append("    .options        = ").Append(HexDump.ShortToHex(field_1_options)).Append("\n");
            sb.Append("         .IsHorizontal = ").Append(HorizontalTextAlignment).Append('\n');
            sb.Append("         .IsVertical   = ").Append(VerticalTextAlignment).Append('\n');
            sb.Append("         .textLocked   = ").Append(IsTextLocked).Append('\n');
            sb.Append("    .textOrientation= ").Append(HexDump.ShortToHex((int) TextOrientation)).Append("\n");
            sb.Append("    .reserved4      = ").Append(HexDump.ShortToHex(field_3_reserved4)).Append("\n");
            sb.Append("    .reserved5      = ").Append(HexDump.ShortToHex(field_4_reserved5)).Append("\n");
            sb.Append("    .reserved6      = ").Append(HexDump.ShortToHex(field_5_reserved6)).Append("\n");
            sb.Append("    .textLength     = ").Append(HexDump.ShortToHex(_text.Length)).Append("\n");
            sb.Append("    .reserved7      = ").Append(HexDump.IntToHex(field_8_reserved7)).Append("\n");

            sb.Append("    .string = ").Append(_text).Append('\n');

            for (int i = 0; i < _text.NumFormattingRuns; i++)
            {
                sb.Append("    .textrun = ").Append(((HSSFRichTextString)_text).GetFontOfFormattingRun(i)).Append('\n');

            }
            sb.Append("[/TXO]\n");
            return sb.ToString();
        }

        public override Object Clone()
        {

            TextObjectRecord rec = new TextObjectRecord();
            rec._text = _text;

            rec.field_1_options = field_1_options;
            rec.field_2_textOrientation = field_2_textOrientation;
            rec.field_3_reserved4 = field_3_reserved4;
            rec.field_4_reserved5 = field_4_reserved5;
            rec.field_5_reserved6 = field_5_reserved6;
            rec.field_8_reserved7 = field_8_reserved7;

            rec._text = _text; // clone needed?

            if (_linkRefPtg != null)
            {
                rec._unknownPreFormulaInt = _unknownPreFormulaInt;
                rec._linkRefPtg = _linkRefPtg.Copy();
                rec._unknownPostFormulaByte = _unknownPostFormulaByte;
            }
            return rec;
        }

    }
}