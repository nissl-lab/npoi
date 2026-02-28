
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


    /**
     * Title:        Print Setup Record
     * Description:  Stores print Setup options -- bogus for HSSF (and marked as such)
     * REFERENCE:  PG 385 Microsoft Excel 97 Developer's Kit (ISBN: 1-57231-498-2)
     * @author Andrew C. Oliver (acoliver at apache dot org)
     * @author Jason Height (jheight at chariot dot net dot au)
     * @version 2.0-pre
     */

    public class PrintSetupRecord
       : StandardRecord
    {

        public const short sid = 0xa1;
        private short field_1_paper_size;
        private short field_2_scale;
        private short field_3_page_start;
        private short field_4_fit_width;
        private short field_5_fit_height;
        private short field_6_options;
        
        private BitField lefttoright =
        BitFieldFactory.GetInstance(0x01);   // print over then down
        private BitField landscape =
        BitFieldFactory.GetInstance(0x02);   // landscape mode
        private BitField validSettings = BitFieldFactory.GetInstance(
        0x04);                // if papersize, scale, resolution, copies, landscape weren't obtained from the print consider them mere bunk
        private BitField nocolor =
        BitFieldFactory.GetInstance(0x08);   // print mono/b&w, colorless
        private BitField draft =
        BitFieldFactory.GetInstance(0x10);   // print draft quality
        private BitField notes =
        BitFieldFactory.GetInstance(0x20);   // print the notes
        private BitField noOrientation =
        BitFieldFactory.GetInstance(0x40);   // the orientation Is not Set
        private BitField usepage =
        BitFieldFactory.GetInstance(0x80);   // use a user Set page no, instead of auto
        private BitField endnote =
        BitFieldFactory.GetInstance(0x200);   // note is printed at the end
        private BitField ierror =
        BitFieldFactory.GetInstance(0xC00);   // printed style of cell errors
       

        private short field_7_hresolution;
        private short field_8_vresolution;
        private double field_9_headermargin;
        private double field_10_footermargin;
        private short field_11_copies;

        public PrintSetupRecord()
        {
            
        }

        /**
         * Constructs a PrintSetup (SetUP) record and Sets its fields appropriately.
         * @param in the RecordInputstream to Read the record from
         */

        public PrintSetupRecord(RecordInputStream in1)
        {
            field_1_paper_size = in1.ReadShort();
            field_2_scale = in1.ReadShort();
            field_3_page_start = in1.ReadShort();
            field_4_fit_width = in1.ReadShort();
            field_5_fit_height = in1.ReadShort();
            field_6_options = in1.ReadShort();
            field_7_hresolution = in1.ReadShort();
            field_8_vresolution = in1.ReadShort();
            field_9_headermargin = in1.ReadDouble();
            field_10_footermargin = in1.ReadDouble();
            field_11_copies = in1.ReadShort();
        }
        public short PaperSize
        {
            get
            {
                return field_1_paper_size;
            }
            set { field_1_paper_size = value; }
        }

        public short Scale
        {
            get{return field_2_scale;}
            set { field_2_scale = value; }
        }

        public short PageStart
        {
            get { return field_3_page_start; }
            set { field_3_page_start = value; }
        }

        public short FitWidth
        {
            get { return field_4_fit_width; }
            set { field_4_fit_width = value; }
        }

        public short FitHeight
        {
            get { return field_5_fit_height; }
            set { field_5_fit_height = value; }
        }

        public short Options
        {
            get { return field_6_options; }
            set { field_6_options = value; }
        }

        // option bitfields
        public bool LeftToRight
        {
            get { return lefttoright.IsSet(field_6_options); }
            set { field_6_options = lefttoright.SetShortBoolean(field_6_options, value); }
        }

        public bool Landscape
        {
            get { return landscape.IsSet(field_6_options); }
            set { field_6_options = landscape.SetShortBoolean(field_6_options, value); }
        }

        public bool ValidSettings
        {
            get { return validSettings.IsSet(field_6_options); }
            set { field_6_options = validSettings.SetShortBoolean(field_6_options, value); }
        }

        public bool NoColor
        {
            get { return nocolor.IsSet(field_6_options); }
            set { field_6_options = nocolor.SetShortBoolean(field_6_options, value); }
        }

        public bool Draft
        {
            get { return draft.IsSet(field_6_options); }
            set { field_6_options = draft.SetShortBoolean(field_6_options, value); }
        }

        public bool Notes
        {
            get
            {
                return notes.IsSet(field_6_options);
            }
            set { field_6_options = notes.SetShortBoolean(field_6_options, value); }
        }

        public bool NoOrientation
        {
            get { return noOrientation.IsSet(field_6_options); }
            set { field_6_options = noOrientation.SetShortBoolean(field_6_options, value); }
        }

        public bool UsePage
        {
            get{return usepage.IsSet(field_6_options);}
            set { field_6_options = usepage.SetShortBoolean(field_6_options, value); }
        }

        // end option bitfields
        public short HResolution
        {
            get { return field_7_hresolution; }
            set { field_7_hresolution = value; }
        }

        public short VResolution
        {
            get { return field_8_vresolution; }
            set { field_8_vresolution = value; }
        }

        public double HeaderMargin
        {
            get { return field_9_headermargin; }
            set { field_9_headermargin = value; }
        }

        public double FooterMargin
        {
            get { return field_10_footermargin; }
            set { field_10_footermargin = value; }
        }

        public bool EndNote
        {
            get { return endnote.IsSet(field_6_options); }
            set { field_6_options = endnote.SetShortBoolean(field_6_options, value); }
        }

        public short CellError
        {
            get { return ierror.GetShortValue(field_6_options); }
            set 
            {
                field_6_options = ierror.SetShortValue(field_6_options, value);
            }
        }

        public short Copies
        {
            get { return field_11_copies; }
            set { field_11_copies = value; }
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[PRINTSetUP]\n");
            buffer.Append("    .papersize      = ").Append(PaperSize)
                .Append("\n");
            buffer.Append("    .scale          = ").Append(Scale)
                .Append("\n");
            buffer.Append("    .pagestart      = ").Append(PageStart)
                .Append("\n");
            buffer.Append("    .fitwidth       = ").Append(FitWidth)
                .Append("\n");
            buffer.Append("    .fitheight      = ").Append(FitHeight)
                .Append("\n");
            buffer.Append("    .options        = ").Append(Options)
                .Append("\n");
            buffer.Append("        .ltor       = ").Append(LeftToRight)
                .Append("\n");
            buffer.Append("        .landscape  = ").Append(Landscape)
                .Append("\n");
            buffer.Append("        .valid      = ").Append(ValidSettings)
                .Append("\n");
            buffer.Append("        .mono       = ").Append(NoColor)
                .Append("\n");
            buffer.Append("        .draft      = ").Append(Draft)
                .Append("\n");
            buffer.Append("        .notes      = ").Append(Notes)
                .Append("\n");
            buffer.Append("        .noOrientat = ").Append(NoOrientation)
                .Append("\n");
            buffer.Append("        .usepage    = ").Append(UsePage)
                .Append("\n");
            buffer.Append("    .hresolution    = ").Append(HResolution)
                .Append("\n");
            buffer.Append("    .vresolution    = ").Append(VResolution)
                .Append("\n");
            buffer.Append("    .headermargin   = ").Append(HeaderMargin)
                .Append("\n");
            buffer.Append("    .footermargin   = ").Append(FooterMargin)
                .Append("\n");
            buffer.Append("    .copies         = ").Append(Copies)
                .Append("\n");
            buffer.Append("[/PRINTSetUP]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteShort(PaperSize);
            out1.WriteShort(Scale);
            out1.WriteShort(PageStart);
            out1.WriteShort(FitWidth);
            out1.WriteShort(FitHeight);
            out1.WriteShort(Options);
            out1.WriteShort(HResolution);
            out1.WriteShort(VResolution);
            out1.WriteDouble(HeaderMargin);
            out1.WriteDouble(FooterMargin);
            out1.WriteShort(Copies);
        }

        protected override int DataSize
        {
            get { return 34; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            PrintSetupRecord rec = new PrintSetupRecord();
            rec.field_1_paper_size = field_1_paper_size;
            rec.field_2_scale = field_2_scale;
            rec.field_3_page_start = field_3_page_start;
            rec.field_4_fit_width = field_4_fit_width;
            rec.field_5_fit_height = field_5_fit_height;
            rec.field_6_options = field_6_options;
            rec.field_7_hresolution = field_7_hresolution;
            rec.field_8_vresolution = field_8_vresolution;
            rec.field_9_headermargin = field_9_headermargin;
            rec.field_10_footermargin = field_10_footermargin;
            rec.field_11_copies = field_11_copies;
            return rec;
        }
    }
}