
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
    using System.Collections;
    using System.IO;
    using System.Text;
    using NPOI.Util;

    /**
     * Title:        Bound Sheet Record (aka BundleSheet) 
     * Description:  Defines a sheet within a workbook.  Basically stores the sheetname
     *               and tells where the Beginning of file record Is within the HSSF
     *               file. 
     * REFERENCE:  PG 291 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Sergei Kozello (sergeikozello at mail.ru)
     * @version 2.0-pre
     */

    public class BoundSheetRecord
           : Record
    {
        public const short sid = 0x85;

        private static BitField hiddenFlag = BitFieldFactory.GetInstance(0x01);
	    private static BitField veryHiddenFlag = BitFieldFactory.GetInstance(0x02);

        private int field_1_position_of_BOF;
        private int field_2_option_flags;
        private int field_4_isMultibyteUnicode;
        private String field_5_sheetname;


        public BoundSheetRecord(String sheetname)
        {
            field_2_option_flags = 0;
            this.Sheetname=sheetname;
        }

        /**
         * Constructs a BoundSheetRecord and Sets its fields appropriately
         *
         * @param in the RecordInputstream to Read the record from
         */

        public BoundSheetRecord(RecordInputStream in1)
        {
            field_1_position_of_BOF = in1.ReadInt();	// bof
            field_2_option_flags = in1.ReadShort();	// flags
            int field_3_sheetname_length = in1.ReadUByte();						// len(str)
            field_4_isMultibyteUnicode = (byte)in1.ReadByte();						// Unicode

            
            if (this.IsMultibyte)
            {
                field_5_sheetname = in1.ReadUnicodeLEString(field_3_sheetname_length);
            }
            else
            {
                field_5_sheetname = in1.ReadCompressedUnicode(field_3_sheetname_length);
            }
        }
        private static void ValidateSheetName(String sheetName)
        {
            if (sheetName == null)
            {
                throw new ArgumentException("sheetName must not be null");
            }
            int len = sheetName.Length;
            if (len < 1)
            {
                throw new ArgumentException("sheetName must not be empty string");
            }
            //if (len > 31)
            //{
            //    throw new ArgumentException("The length of sheetName must not be longer than 31");
            //}
            for (int i = 0; i < len; i++)
            {
                char ch = sheetName[i];
                switch (ch)
                {
                    case '/':
                    case '\\':
                    case '?':
                    case '*':
                    case ']':
                    case '[':
                        break;
                    default:
                        // all other chars OK
                        continue;
                }
                throw new ArgumentException("Invalid char (" + ch
                        + ") found at index (" + i + ") in sheet name '" + sheetName + "'");
            }
        }

        /**
         * Get the offset in bytes of the Beginning of File Marker within the HSSF Stream part of the POIFS file
         *
         * @return offset in bytes
         */

        public int PositionOfBof
        {
            get { return field_1_position_of_BOF; }
            set { field_1_position_of_BOF = value; }
        }

        /**
         * Is the sheet very hidden? Different from (normal) hidden 
         */
        public bool IsVeryHidden
        {
            get
            {
                return veryHiddenFlag.IsSet(field_2_option_flags);
            }
            set 
            {
                field_2_option_flags = veryHiddenFlag.SetBoolean(field_2_option_flags, value);
            }
        }

        /**
         * Get the sheetname for this sheet.  (this appears in the tabs at the bottom)
         * @return sheetname the name of the sheet
         */

        public String Sheetname
        {
            get { return field_5_sheetname; }
            set
            {
                ValidateSheetName(value);
                field_5_sheetname = value;
                field_4_isMultibyteUnicode = (StringUtil.HasMultibyte(value) ? (byte)1 : (byte)0);

            }
        }

        private bool IsMultibyte
        {
            get
            {
                return (field_4_isMultibyteUnicode & 0x01) != 0;
            }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[BOUNDSHEET]\n");
            buffer.Append("    .bof             = ")
                    .Append(HexDump.IntToHex(PositionOfBof)).Append("\n");
            buffer.Append("    .optionflags     = ")
                    .Append(HexDump.ShortToHex(field_2_option_flags)).Append("\n");
            buffer.Append("    .sheetname Length= ")
                    .Append(HexDump.ByteToHex(field_4_isMultibyteUnicode)).Append("\n");
            buffer.Append("    .sheetname       = ").Append(Sheetname)
                    .Append("\n");
            buffer.Append("[/BOUNDSHEET]\n");
            return buffer.ToString();
        }

        private int DataSize
        {
            get
            {
                return 8 + field_5_sheetname.Length * (IsMultibyte ? 2 : 1);
            }
        }

        public override int Serialize(int offset, byte [] data)
        {
            LittleEndian.PutUShort(data, 0 + offset, sid);
            LittleEndian.PutUShort(data, 2 + offset, DataSize);
            LittleEndian.PutInt(data, 4 + offset, PositionOfBof);
            LittleEndian.PutUShort(data, 8 + offset, field_2_option_flags);

            String name = field_5_sheetname;
            LittleEndian.PutByte(data, 10 + offset, name.Length);
            LittleEndian.PutByte(data, 11 + offset, field_4_isMultibyteUnicode);

            if (IsMultibyte)
                StringUtil.PutUnicodeLE(Sheetname, data, 12 + offset);
            else
                StringUtil.PutCompressedUnicode(Sheetname, data, 12 + offset);


            return RecordSize;

            /*
            byte[] fake = new byte[] {	(byte)0x85, 0x00, 			// sid
                                            0x1a, 0x00, 			// Length
                                            0x3C, 0x09, 0x00, 0x00, // bof
                                            0x00, 0x00, 			// flags
                                            0x09, 					// len( str )
                                            0x01, 					// Unicode
                                            // <str>
                                            0x21, 0x04, 0x42, 0x04, 0x40, 0x04, 0x30, 0x04, 0x3D,
                                            0x04, 0x38, 0x04, 0x47, 0x04, 0x3A, 0x04, 0x30, 0x04
                                            // </str>
                                        };

                                        sid + len + bof + flags + len(str) + Unicode +   str
                                         2  +  2  +  4  +   2   +    1     +    1    + len(str)

            Array.Copy( fake, 0, data, offset, fake.Length );

            return fake.Length;
            */
        }

        public override int RecordSize
        {
            get { return 4 + DataSize; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public bool IsHidden
        {
            get
            {
                return hiddenFlag.IsSet(field_2_option_flags);
            }
            set
            {
                field_2_option_flags = hiddenFlag.SetBoolean(field_2_option_flags, value);
            }
        }

        	/**
	     * Converts a List of {@link BoundSheetRecord}s to an array and sorts by the position of their
	     * BOFs.
	     */
	    public static BoundSheetRecord[] OrderByBofPosition(ArrayList boundSheetRecords) 
        {
		    
		    BoundSheetRecord[] bsrs = 
                (BoundSheetRecord[])boundSheetRecords.ToArray(typeof(BoundSheetRecord));
		    Array.Sort(bsrs, new BOFComparator());
	 	    return bsrs;
	    }



	    private class BOFComparator:IComparer
        {
		    public int Compare(BoundSheetRecord bsr1, BoundSheetRecord bsr2) {
			    return bsr1.PositionOfBof - bsr2.PositionOfBof;
		    }

            #region IComparer Members

            int IComparer.Compare(object x, object y)
            {
                return Compare((BoundSheetRecord)x, (BoundSheetRecord)y);
            }

            #endregion
        };
    }
}