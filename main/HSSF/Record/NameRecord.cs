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

namespace NPOI.HSSF.Record
{

    using System;
    using System.Text;

    using NPOI.Util;
    using NPOI.HSSF.Record;
    using SSFormula=NPOI.SS.Formula;
    using NPOI.HSSF.Record.Cont;
    using NPOI.SS.Formula.PTG;

    /**
     * Title:        Name Record (aka Named Range) 
     * Description:  Defines a named range within a workbook. 
     * REFERENCE:  
     * @author Libin Roman (Vista Portal LDT. Developer)
     * @author  Sergei Kozello (sergeikozello at mail.ru)
     * @author Glen Stampoultzis (glens at apache.org)
     * @version 1.0-pre
     */
    public class NameRecord : ContinuableRecord
    {
        private enum Option:short {
		    OPT_HIDDEN_NAME =   0x0001,
		    OPT_FUNCTION_NAME = 0x0002,
            OPT_COMMAND_NAME = 0x0004,
            OPT_MACRO = 0x0008,
            OPT_COMPLEX = 0x0010,
            OPT_BUILTIN = 0x0020,
            OPT_BINDATA = 0x1000,
	    }

        public static bool IsFormula(int optValue) {
			return (optValue & 0x0F) == 0;
		}

        /**
         */
        public const short sid = 0x18; //Docs says that it Is 0x218

        /**Included for completeness sake, not implemented
           */
        public const byte BUILTIN_CONSOLIDATE_AREA = (byte)0;

        /**Included for completeness sake, not implemented
         */
        public const byte BUILTIN_AUTO_OPEN = (byte)1;

        /**Included for completeness sake, not implemented
         */
        public const byte BUILTIN_AUTO_CLOSE = (byte)2;


        public const byte BUILTIN_EXTRACT = (byte)3;
        /**Included for completeness sake, not implemented
         */
        public const byte BUILTIN_DATABASE = (byte)4;

        /**Included for completeness sake, not implemented
         */
        public const byte BUILTIN_CRITERIA = (byte)5;

        public const byte BUILTIN_PRINT_AREA = (byte)6;
        public const byte BUILTIN_PRINT_TITLE = (byte)7;

        /**Included for completeness sake, not implemented
         */
        public const byte BUILTIN_RECORDER = (byte)8;

        /**Included for completeness sake, not implemented
         */
        public const byte BUILTIN_DATA_FORM = (byte)9;

        /**Included for completeness sake, not implemented
         */

        public const byte BUILTIN_AUTO_ACTIVATE = (byte)10;

        /**Included for completeness sake, not implemented
         */

        public const byte BUILTIN_AUTO_DEACTIVATE = (byte)11;

        /**Included for completeness sake, not implemented
         */
        public const byte BUILTIN_SHEET_TITLE = (byte)12;

        public const byte  BUILTIN_FILTER_DB             = 13;

        //public const short OPT_HIDDEN_NAME = (short)0x0001;
        //public const short OPT_FUNCTION_NAME = (short)0x0002;
        //public const short OPT_COMMAND_NAME = (short)0x0004;
        //public const short OPT_MACRO = (short)0x0008;
        //public const short OPT_COMPLEX = (short)0x0010;
        //public const short OPT_BUILTIN = (short)0x0020;
        //public const short OPT_BINDATA = (short)0x1000;


        private short field_1_option_flag;
        private byte field_2_keyboard_shortcut;
        //private byte field_3_Length_name_text;
        //private short field_4_Length_name_definition;
        //private short field_5_index_to_sheet;     // Unused: see field_6
        //private short field_6_Equals_to_index_to_sheet;
        /** One-based extern index of sheet (resolved via LinkTable). Zero if this is a global name  */
        private short field_5_externSheetIndex_plus1;
        /** the one based sheet number.  */
        private int field_6_sheetNumber;
        //private byte field_7_Length_custom_menu;
        //private byte field_8_Length_description_text;
        //private byte field_9_Length_help_topic_text;
        //private byte field_10_Length_status_bar_text;
        //private byte field_11_compressed_unicode_flag;   // not documented
        private bool field_11_nameIsMultibyte;
        private byte field_12_built_in_code;
        private String field_12_name_text;
        private SSFormula.Formula field_13_name_definition;
        private String field_14_custom_menu_text;
        private String field_15_description_text;
        private String field_16_help_topic_text;
        private String field_17_status_bar_text;


        /** Creates new NameRecord */
        public NameRecord()
        {
            field_13_name_definition = SSFormula.Formula.Create(Ptg.EMPTY_PTG_ARRAY);

            field_12_name_text = "";
            field_14_custom_menu_text = "";
            field_15_description_text = "";
            field_16_help_topic_text = "";
            field_17_status_bar_text = "";
        }
        protected int DataSize
        {
            get {
                return 13   // 3 shorts + 7 bytes
                        + NameRawSize
                        + field_14_custom_menu_text.Length
                        + field_15_description_text.Length
                        + field_16_help_topic_text.Length
                        + field_17_status_bar_text.Length
                        + field_13_name_definition.EncodedSize;
            }
        }
        /**
         * Constructs a Name record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */
        public NameRecord(RecordInputStream ris)
        {
            byte[] remainder = ris.ReadAllContinuedRemainder();
            ILittleEndianInput in1 = new LittleEndianByteArrayInputStream(remainder);
            field_1_option_flag                 = in1.ReadShort();
		    field_2_keyboard_shortcut           = (byte)in1.ReadByte();
		    int field_3_length_name_text        = in1.ReadByte();
		    int field_4_length_name_definition  = in1.ReadShort();
		    field_5_externSheetIndex_plus1      = in1.ReadShort();
		    field_6_sheetNumber                 = in1.ReadUShort();
		    int field_7_length_custom_menu      = in1.ReadUByte();
		    int field_8_length_description_text = in1.ReadUByte();
		    int field_9_length_help_topic_text  = in1.ReadUByte();
		    int field_10_length_status_bar_text = in1.ReadUByte();

		    //store the name in byte form if it's a built-in name
		    field_11_nameIsMultibyte = (in1.ReadByte() != 0);
		    if (IsBuiltInName) {
			    field_12_built_in_code = (byte)in1.ReadByte();
		    } else {
			    if (field_11_nameIsMultibyte) {
                    field_12_name_text = StringUtil.ReadUnicodeLE(in1, field_3_length_name_text);
			    } else {
                    field_12_name_text = StringUtil.ReadCompressedUnicode(in1, field_3_length_name_text);
			    }
		    }
          int nBytesAvailable = in1.Available() - (field_7_length_custom_menu
				+ field_8_length_description_text + field_9_length_help_topic_text + field_10_length_status_bar_text);
		    field_13_name_definition = SSFormula.Formula.Read(field_4_length_name_definition, in1, nBytesAvailable);

		    //Who says that this can only ever be compressed unicode???
            field_14_custom_menu_text = StringUtil.ReadCompressedUnicode(in1, field_7_length_custom_menu);
            field_15_description_text = StringUtil.ReadCompressedUnicode(in1, field_8_length_description_text);
            field_16_help_topic_text = StringUtil.ReadCompressedUnicode(in1, field_9_length_help_topic_text);
            field_17_status_bar_text = StringUtil.ReadCompressedUnicode(in1, field_10_length_status_bar_text);

        }

        /**
         * Constructor to Create a built-in named region
         * @param builtin Built-in byte representation for the name record, use the public constants
         * @param index 
         */
        public NameRecord(byte builtin, int sheetNumber)
            : this()
        {

            field_12_built_in_code = builtin;
            OptionFlag=(short)(field_1_option_flag | (short)Option.OPT_BUILTIN);
            field_6_sheetNumber = sheetNumber; //the extern sheets are set through references
        }

        /**
         * @return function Group
         * @see FnGroupCountRecord
         */
        public byte FnGroup
        {
            get
            {
                int masked = field_1_option_flag & 0x0fc0;
                return (byte)(masked >> 4);
            }
        }

        /** Gets the option flag
         * @return option flag
         */
        public short OptionFlag
        {
            get { return field_1_option_flag; }
            set { field_1_option_flag = value; }
        }

        /** returns the keyboard shortcut
         * @return keyboard shortcut
         */
        public byte KeyboardShortcut
        {
            get { return field_2_keyboard_shortcut; }
            set { field_2_keyboard_shortcut = value; }
        }

        ///** 
        // * Gets the name Length, in Chars
        // * @return name Length
        // */
        public byte NameTextLength
        {
            get
            {
                if (IsBuiltInName)
                {
                    return 1;
                }
                return (byte)field_12_name_text.Length;
            }
        }

        private int NameRawSize
        {
            get
            {
                if (IsBuiltInName)
                {
                    return 1;
                }
                int nChars = field_12_name_text.Length;
                if (field_11_nameIsMultibyte)
                {
                    return 2 * nChars;
                }
                return nChars;
            }
        }

        /**
	 * Indicates that the defined name refers to a user-defined function.
	 * This attribute is used when there is an add-in or other code project associated with the file.
	 *
	 * @param function <c>true</c> indicates the name refers to a function.
	 */
        public void SetFunction(bool function)
        {
            if (function)
            {
                field_1_option_flag |= (short)Option.OPT_FUNCTION_NAME;
            }
            else
            {
                field_1_option_flag &= (short)(~Option.OPT_FUNCTION_NAME);
            }
        }
        /**
 * @return <c>true</c> if name has a formula (named range or defined value)
 */
        public bool HasFormula
        {
            get
            {
                return IsFormula(field_1_option_flag) && field_13_name_definition.EncodedTokenSize > 0;
            }
        }
        /**
         * @return true if name Is hidden
         */
        public bool IsHiddenName
        {
            get { return (field_1_option_flag & (short)Option.OPT_HIDDEN_NAME) != 0; }
            set 
            {
                if (value)
                {
                    field_1_option_flag |= (short)Option.OPT_HIDDEN_NAME;
                }
                else
                {
                    field_1_option_flag &= (short)(~Option.OPT_HIDDEN_NAME);
                }
            }
        }


        /**
         * @return true if name Is a function
         */
        public bool IsFunctionName
        {
            get { return (field_1_option_flag & (short)Option.OPT_FUNCTION_NAME) != 0; }
            set
            {
                if (value)
                {
                    field_1_option_flag |= (short)Option.OPT_FUNCTION_NAME;
                }
                else
                {
                    field_1_option_flag &= (~(short)Option.OPT_FUNCTION_NAME);
                }
            }
        }

        /**
         * @return true if name Is a command
         */
        public bool IsCommandName
        {
            get { return (field_1_option_flag & (short)Option.OPT_COMMAND_NAME) != 0; }
        }

        /**
         * @return true if function macro or command macro
         */
        public bool IsMacro
        {
            get { return (field_1_option_flag & (short)Option.OPT_MACRO) != 0; }
        }

        /**
         * @return true if array formula or user defined
         */
        public bool IsComplexFunction
        {
            get { return (field_1_option_flag & (short)Option.OPT_COMPLEX) != 0; }
        }


        /**Convenience Function to determine if the name Is a built-in name
         */
        public bool IsBuiltInName
        {
            get { return ((this.OptionFlag & (short)Option.OPT_BUILTIN) != 0); }
        }


        /** Gets the name
         * @return name
         */
        public String NameText
        {
            get
            {
                return this.IsBuiltInName ? this.TranslateBuiltInName(this.BuiltInName) : field_12_name_text;
            }
            set
            {
                field_12_name_text = value;
                field_11_nameIsMultibyte = StringUtil.HasMultibyte(value);
            }
        }

        /** Gets the Built In Name
         * @return the built in Name
         */
        public byte BuiltInName
        {
            get { return this.field_12_built_in_code; }
        }


        /** Gets the definition, reference (Formula)
         * @return definition -- can be null if we cant Parse ptgs
         */
        public Ptg[] NameDefinition
        {
            get { return field_13_name_definition.Tokens; }
            set { field_13_name_definition = SSFormula.Formula.Create(value); }
        }
        /** Get the custom menu text
         * @return custom menu text
         */
        public String CustomMenuText
        {
            get { return field_14_custom_menu_text; }
            set { field_14_custom_menu_text = value; }
        }

        /** Gets the description text
         * @return description text
         */
        public String DescriptionText
        {
            get { return field_15_description_text; }
            set { field_15_description_text = value; }
        }

        /** Get the help topic text
         * @return gelp topic text
         */
        public String HelpTopicText
        {
            get { return field_16_help_topic_text; }
            set { field_16_help_topic_text = value; }
        }

        /** Gets the status bar text
         * @return status bar text
         */
        public String StatusBarText
        {
            get { return field_17_status_bar_text; }
            set { field_17_status_bar_text = value; }
        }
        /**
 * For named ranges, and built-in names
 * @return the 1-based sheet number. 
 */
        public int SheetNumber
        {
            get
            {
                return field_6_sheetNumber;
            }
            set { field_6_sheetNumber = value; }
        }
        /**
         * called by the class that Is responsible for writing this sucker.
         * Subclasses should implement this so that their data Is passed back in a
         * @param offset to begin writing at
         * @param data byte array containing instance data
         * @return number of bytes written
         */
        protected override void Serialize(ContinuableRecordOutput out1)
        {
            int field_7_length_custom_menu = field_14_custom_menu_text.Length;
            int field_8_length_description_text = field_15_description_text.Length;
            int field_9_length_help_topic_text = field_16_help_topic_text.Length;
            int field_10_length_status_bar_text = field_17_status_bar_text.Length;
            //int rawNameSize = NameRawSize;

            // size defined below
            out1.WriteShort(OptionFlag);
            out1.WriteByte(KeyboardShortcut);
            out1.WriteByte(NameTextLength);
            // Note - formula size is not immediately before encoded formula, and does not include any array constant data
            out1.WriteShort(field_13_name_definition.EncodedTokenSize);
            out1.WriteShort(field_5_externSheetIndex_plus1);
            out1.WriteShort(field_6_sheetNumber);
            out1.WriteByte(field_7_length_custom_menu);
            out1.WriteByte(field_8_length_description_text);
            out1.WriteByte(field_9_length_help_topic_text);
            out1.WriteByte(field_10_length_status_bar_text);
            out1.WriteByte(field_11_nameIsMultibyte ? 1 : 0);

            if (IsBuiltInName)
            {
                out1.WriteByte(field_12_built_in_code);
            }
            else
            {
                String nameText = field_12_name_text;
                if (field_11_nameIsMultibyte)
                {
                    StringUtil.PutUnicodeLE(nameText,out1);
                }
                else
                {
                    StringUtil.PutCompressedUnicode(nameText, out1);
                }
            }
            field_13_name_definition.SerializeTokens(out1);
            field_13_name_definition.SerializeArrayConstantData(out1);

            StringUtil.PutCompressedUnicode(CustomMenuText,out1);
            StringUtil.PutCompressedUnicode(DescriptionText, out1);
            StringUtil.PutCompressedUnicode(HelpTopicText, out1);
            StringUtil.PutCompressedUnicode(StatusBarText, out1);
        }

        /** Gets the extern sheet number
         * @return extern sheet index
         */
        public int ExternSheetNumber
        {
            get
            {
                if (field_13_name_definition.EncodedSize < 1)
                {
                    return 0;
                }
                Ptg ptg = field_13_name_definition.Tokens[0];

                if (ptg.GetType() == typeof(Area3DPtg))
                {
                    return ((Area3DPtg)ptg).ExternSheetIndex;

                }
                else if (ptg.GetType() == typeof(Ref3DPtg))
                {
                    return ((Ref3DPtg)ptg).ExternSheetIndex;
                }

                return 0;
            }
        }


        private Ptg CreateNewPtg()
        {
            return new Area3DPtg("A1:A1", 0); // TODO - change to not be partially initialised
        }

        /**
         * return the non static version of the id for this record.
         */
        public override short Sid
        {
            get { return sid; }
        }
        /*
          20 00 
          00 
          01 
          1A 00 // sz = 0x1A = 26
          00 00 
          01 00 
          00 
          00 
          00 
          00 
          00 // Unicode flag
          07 // name
      
          29 17 00 3B 00 00 00 00 FF FF 00 00 02 00 3B 00 //{ 26
          00 07 00 07 00 00 00 FF 00 10                   //  }
      
      
      
          20 00 
          00 
          01 
          0B 00 // sz = 0xB = 11
          00 00 
          01 00 
          00 
          00 
          00 
          00 
          00 // Unicode flag
          07 // name
      
          3B 00 00 07 00 07 00 00 00 FF 00   // { 11 }
      */
        /*
          18, 00, 
          1B, 00, 
      
          20, 00, 
          00, 
          01, 
          0B, 00, 
          00, 
          00, 
          00, 
          00, 
          00, 
          07, 
          3B 00 00 07 00 07 00 00 00 FF 00 ]     
         */

        /**
         * @see Object#ToString()
         */
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[NAME]\n");
            buffer.Append("    .option flags         = ").Append(HexDump.ToHex(field_1_option_flag))
                .Append("\n");
            buffer.Append("    .keyboard shortcut    = ").Append(HexDump.ToHex(field_2_keyboard_shortcut))
                .Append("\n");
            buffer.Append("    .Length of the name   = ").Append(NameTextLength)
                .Append("\n");
            buffer.Append("    .extSheetIx(1-based, 0=Global)= ").Append(field_5_externSheetIndex_plus1)
                .Append("\n");
            buffer.Append("    .sheetTabIx           = ").Append(field_6_sheetNumber)
                .Append("\n");
            buffer.Append("    .Length of menu text (Char count)        = ").Append(field_14_custom_menu_text.Length)
                .Append("\n");
            buffer.Append("    .Length of description text (Char count) = ").Append(field_15_description_text.Length)
                .Append("\n");
            buffer.Append("    .Length of help topic text (Char count)  = ").Append(field_16_help_topic_text.Length)
                .Append("\n");
            buffer.Append("    .Length of status bar text (Char count)  = ").Append(field_17_status_bar_text.Length)
                .Append("\n");
            buffer.Append("    .Name (Unicode flag)  = ").Append(field_11_nameIsMultibyte)
                .Append("\n");
            buffer.Append("    .Name (Unicode text)  = ").Append(NameText)
                .Append("\n");
            Ptg[] ptgs = field_13_name_definition.Tokens;
            buffer.AppendLine("    .Formula (nTokens=" + ptgs.Length + "):");
            for (int i = 0; i < ptgs.Length; i++)
            {
                Ptg ptg = ptgs[i];
                buffer.Append("       " + ptg.ToString()).Append(ptg.RVAType).Append("\n");
            }

            buffer.Append("    .Menu text (Unicode string without Length field)        = ").Append(field_14_custom_menu_text)
                .Append("\n");
            buffer.Append("    .Description text (Unicode string without Length field) = ").Append(field_15_description_text)
                .Append("\n");
            buffer.Append("    .Help topic text (Unicode string without Length field)  = ").Append(field_16_help_topic_text)
                .Append("\n");
            buffer.Append("    .Status bar text (Unicode string without Length field)  = ").Append(field_17_status_bar_text)
                .Append("\n");
            buffer.Append("[/NAME]\n");

            return buffer.ToString();
        }

        /**Creates a human Readable name for built in types
         * @return Unknown if the built-in name cannot be translated
         */
        protected String TranslateBuiltInName(byte name)
        {
            switch (name)
            {
                case NameRecord.BUILTIN_AUTO_ACTIVATE: return "Auto_Activate";
                case NameRecord.BUILTIN_AUTO_CLOSE: return "Auto_Close";
                case NameRecord.BUILTIN_AUTO_DEACTIVATE: return "Auto_Deactivate";
                case NameRecord.BUILTIN_AUTO_OPEN: return "Auto_Open";
                case NameRecord.BUILTIN_CONSOLIDATE_AREA: return "Consolidate_Area";
                case NameRecord.BUILTIN_CRITERIA: return "Criteria";
                case NameRecord.BUILTIN_DATABASE: return "Database";
                case NameRecord.BUILTIN_DATA_FORM: return "Data_Form";
                case NameRecord.BUILTIN_PRINT_AREA: return "Print_Area";
                case NameRecord.BUILTIN_PRINT_TITLE: return "Print_Titles";
                case NameRecord.BUILTIN_RECORDER: return "Recorder";
                case NameRecord.BUILTIN_SHEET_TITLE: return "Sheet_Title";
                case NameRecord.BUILTIN_FILTER_DB: return "_FilterDatabase";
                case NameRecord.BUILTIN_EXTRACT: return "Extract";
            }

            return "Unknown";
        }
    }
}