
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

    /// <summary>
    /// WSBool Record. Stores workbook Settings  (aka its a big "everything we didn't put somewhere else")
    /// </summary>
    /// <remarks>REFERENCE:  PG 425 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
    /// </remarks>
    public class WSBoolRecord
       : StandardRecord
    {
        public const short sid = 0x81;
        private byte field_1_wsbool;         // crappy names are because this is really one big short field (2byte)
        private byte field_2_wsbool;         // but the docs inconsistantly use it as 2 seperate bytes

        // I decided to be consistant in this way.
        static private BitField autobreaks =
            BitFieldFactory.GetInstance(0x01);                               // are automatic page breaks visible

        // bits 1 to 3 Unused
        static private BitField dialog = BitFieldFactory.GetInstance(0x10);                               // is sheet dialog sheet
        static private BitField applystyles = BitFieldFactory.GetInstance(0x20);                               // whether to apply automatic styles to outlines
        static private BitField rowsumsbelow = BitFieldFactory.GetInstance(0x40);                                            // whether summary rows will appear below detail in outlines
        static private BitField rowsumsright = BitFieldFactory.GetInstance(0x80);                                            // whether summary rows will appear right of the detail in outlines
        static private BitField fittopage = BitFieldFactory.GetInstance(0x01);                               // whether to fit stuff to the page

        // bit 2 reserved
        static private BitField Displayguts = BitFieldFactory.GetInstance(0x06);                                            // whether to Display outline symbols (in the gutters)

        // bits 4-5 reserved
        static private BitField alternateexpression = BitFieldFactory.GetInstance(0x40); // whether to use alternate expression eval

        static private BitField alternateformula = BitFieldFactory.GetInstance(0x80); // whether to use alternate formula entry


        public WSBoolRecord()
        {
        }


        /// <summary>
        /// Constructs a WSBool record and Sets its fields appropriately.
        /// </summary>
        /// <param name="in1">the RecordInputstream to Read the record from</param>
        public WSBoolRecord(RecordInputStream in1)
        {
            byte[] data = in1.ReadRemainder();
            field_1_wsbool =
                data[0];
            field_2_wsbool =
                data[1];
        }

        /// <summary>
        /// Get first byte (see bit Getters)
        /// </summary>
        public byte WSBool1
        {
            get
            {
                return field_1_wsbool;
            }
            set { field_1_wsbool = value; }
        }

        // bool1 bitfields

        /// <summary>
        /// Whether to show automatic page breaks or not
        /// </summary>
        public bool Autobreaks
        {
            get
            {
                return autobreaks.IsSet(field_1_wsbool);
            }
            set { field_1_wsbool = autobreaks.SetByteBoolean(field_1_wsbool, value); }
        }

        /// <summary>
        /// Whether sheet is a dialog sheet or not
        /// </summary>
        public bool Dialog
        {
            get
            {
                return dialog.IsSet(field_1_wsbool);
            }
            set { field_1_wsbool = dialog.SetByteBoolean(field_1_wsbool, value); }
        }

        /// <summary>
        /// Get if row summaries appear below detail in the outline
        /// </summary>
        public bool RowSumsBelow
        {
            get
            {
                return rowsumsbelow.IsSet(field_1_wsbool);
            }
            set { field_1_wsbool = rowsumsbelow.SetByteBoolean(field_1_wsbool, value); }
        }

        /// <summary>
        /// Get if col summaries appear right of the detail in the outline
        /// </summary>
        public bool RowSumsRight
        {
            get
            {
                return rowsumsright.IsSet(field_1_wsbool);
            }
            set { field_1_wsbool = rowsumsright.SetByteBoolean(field_1_wsbool, value); }
        }


        /// <summary>
        /// Get the second byte (see bit Getters)
        /// </summary>
        public byte WSBool2
        {
            get
            {
                return field_2_wsbool;
            }
            set { field_2_wsbool = value; }
        }


        /// <summary>
        /// fit to page option is on
        /// </summary>
        public bool FitToPage
        {
            get
            {
                return fittopage.IsSet(field_2_wsbool);
            }
            set { field_2_wsbool = fittopage.SetByteBoolean(field_2_wsbool, value); }
        }

        /// <summary>
        /// Whether to display the guts or not
        /// </summary>
        public bool DisplayGuts
        {
            get
            {
                return Displayguts.IsSet(field_2_wsbool);
            }
            set { field_2_wsbool = Displayguts.SetByteBoolean(field_2_wsbool, value); }
        }

        /// <summary>
        /// whether alternate expression evaluation is on
        /// </summary>
        public bool AlternateExpression
        {
            get
            {
                return alternateexpression.IsSet(field_2_wsbool);
            }
            set
            {
                field_2_wsbool = alternateexpression.SetByteBoolean(field_2_wsbool,
                    value);
            }
        }

        /// <summary>
        /// whether alternative formula entry is on
        /// </summary>
        public bool AlternateFormula
        {
            get
            {
                return alternateformula.IsSet(field_2_wsbool);
            }
            set
            {
                field_2_wsbool = alternateformula.SetByteBoolean(field_2_wsbool,
                    value);
            }
        }

        public override string ToString()
        {
            return "[WSBOOL]\n" +
                "    .wsbool1        = " + HexDump.ToHex(WSBool1) + "\n" +
                "        .autobreaks = " + Autobreaks + "\n" +
                "        .dialog     = " + Dialog + "\n" +
                "        .rowsumsbelw= " + RowSumsBelow + "\n" +
                "        .rowsumsrigt= " + RowSumsRight + "\n" +
                "    .wsbool2        = " + HexDump.ToHex(WSBool2) + "\n" +
                "        .fittopage  = " + FitToPage + "\n" +
                "        .displayguts= " + DisplayGuts + "\n" +
                "        .alternateex= " + AlternateExpression + "\n" +
                "        .alternatefo= " + AlternateFormula + "\n" +
                "[/WSBOOL]\n";
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte(WSBool1);
            out1.WriteByte(WSBool2);
        }

        protected override int DataSize
        {
            get { return 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            WSBoolRecord rec = new WSBoolRecord();
            rec.field_1_wsbool = field_1_wsbool;
            rec.field_2_wsbool = field_2_wsbool;
            return rec;
        }
    }
}