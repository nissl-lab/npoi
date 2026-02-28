
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


    /*
     * The text record is used to define text stored on a chart.
     * NOTE: This source is automatically generated please do not modify this file.  Either subclass or
     *       Remove the record in src/records/definitions.

     * @author Glen Stampoultzis (glens at apache.org)
     */
    //
    /// <summary>
    /// Section [2.4.324]. The Text record specifies the properties of an attached label and specifies the beginning of 
    /// a collection of records as defined by the chart sheet substream ABNF. This collection of records specifies an attached label.
    /// </summary>
    public class TextRecord : StandardRecord
    {
        public const short sid = 0x1025;
        private byte field_1_horizontalAlignment;
        /// <summary>
        /// Left-alignment if iReadingOrder specifies left-to-right reading order; otherwise, right-alignment
        /// </summary>
        public const byte HORIZONTAL_ALIGNMENT_LEFT = 1;
        /// <summary>
        /// Center-alignment
        /// </summary>
        public const byte HORIZONTAL_ALIGNMENT_CENTER = 2;
        /// <summary>
        /// Right-alignment if iReadingOrder specifies left-to-right reading order; otherwise, left-alignment
        /// </summary>
        public const byte HORIZONTAL_ALIGNMENT_BOTTOM = 3;
        /// <summary>
        /// Justify-alignment
        /// </summary>
        public const byte HORIZONTAL_ALIGNMENT_JUSTIFY = 4;
        /// <summary>
        /// distributed alignment
        /// </summary>
        public const byte HORIZONTAL_ALIGNMENT_DISTRIBUTED = 7;

        private byte field_2_verticalAlignment;
        public const byte VERTICAL_ALIGNMENT_TOP = 1;
        public const byte VERTICAL_ALIGNMENT_CENTER = 2;
        public const byte VERTICAL_ALIGNMENT_BOTTOM = 3;
        public const byte VERTICAL_ALIGNMENT_JUSTIFY = 4;
        /// <summary>
        /// distributed alignment
        /// </summary>
        public const byte VERTICAL_ALIGNMENT_DISTRIBUTED = 7;

        private short field_3_DisplayMode;
        /// <summary>
        /// Transparent background
        /// </summary>
        public const short DISPLAY_MODE_TRANSPARENT = 1;
        /// <summary>
        /// Opaque background
        /// </summary>
        public const short DISPLAY_MODE_OPAQUE = 2;
        private int field_4_rgbColor;
        private int field_5_x;
        private int field_6_y;
        private int field_7_width;
        private int field_8_height;
        private short field_9_options1;
        private BitField autoColor = BitFieldFactory.GetInstance(0x1);
        private BitField showKey = BitFieldFactory.GetInstance(0x2);
        private BitField showValue = BitFieldFactory.GetInstance(0x4);
        //private BitField vertical = BitFieldFactory.GetInstance(0x8);  //unused

        private BitField autoText = BitFieldFactory.GetInstance(0x10);
        private BitField generated = BitFieldFactory.GetInstance(0x20);
        private BitField autoLabelDeleted = BitFieldFactory.GetInstance(0x40);
        private BitField autoBackground = BitFieldFactory.GetInstance(0x80);
        //private BitField rotation = BitFieldFactory.GetInstance(0x0700); //unused2
        public const short ROTATION_NONE = 0;
        public const short ROTATION_TOP_TO_BOTTOM = 1;
        public const short ROTATION_ROTATED_90_DEGREES = 2;
        public const short ROTATION_ROTATED_90_DEGREES_CLOCKWISE = 3;
        private BitField showCategoryLabelAsPercentage = BitFieldFactory.GetInstance(0x800);
        private BitField showValueAsPercentage = BitFieldFactory.GetInstance(0x1000);
        private BitField showBubbleSizes = BitFieldFactory.GetInstance(0x2000);
        private BitField showLabel = BitFieldFactory.GetInstance(0x4000);
        private short field_10_IndexOfColorValue;
        private short field_11_options2;
        private BitField dataLabelPlacement = BitFieldFactory.GetInstance(0x000F);
        public const short DATA_LABEL_PLACEMENT_CHART_DEPENDENT = 0;
        public const short DATA_LABEL_PLACEMENT_OUTSIDE = 1;
        public const short DATA_LABEL_PLACEMENT_INSIDE = 2;
        public const short DATA_LABEL_PLACEMENT_CENTER = 3;
        public const short DATA_LABEL_PLACEMENT_AXIS = 4;
        public const short DATA_LABEL_PLACEMENT_ABOVE = 5;
        public const short DATA_LABEL_PLACEMENT_BELOW = 6;
        public const short DATA_LABEL_PLACEMENT_LEFT = 7;
        public const short DATA_LABEL_PLACEMENT_RIGHT = 8;
        public const short DATA_LABEL_PLACEMENT_AUTO = 9;
        public const short DATA_LABEL_PLACEMENT_USER_MOVED = 10;

        private BitField readingOrder = BitFieldFactory.GetInstance(0xC000);
        public const short READING_ORDER_CONTEXT = 0;
        public const short READING_ORDER_LTR = 1;
        public const short READING_ORDER_RTL = 2;

        private short field_12_textRotation;


        public TextRecord()
        {

        }

        /**
         * Constructs a Text record and Sets its fields appropriately.
         *
         * @param in the RecordInputstream to Read the record from
         */

        public TextRecord(RecordInputStream in1)
        {
            field_1_horizontalAlignment = (byte)in1.ReadByte();
            field_2_verticalAlignment = (byte)in1.ReadByte();
            field_3_DisplayMode = in1.ReadShort();
            field_4_rgbColor = in1.ReadInt();
            field_5_x = in1.ReadInt();
            field_6_y = in1.ReadInt();
            field_7_width = in1.ReadInt();
            field_8_height = in1.ReadInt();
            field_9_options1 = in1.ReadShort();
            field_10_IndexOfColorValue = in1.ReadShort();
            field_11_options2 = in1.ReadShort();
            field_12_textRotation = in1.ReadShort();
        }

        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append("[TEXT]\n");
            buffer.Append("    .horizontalAlignment  = ")
                .Append("0x").Append(HexDump.ToHex(HorizontalAlignment))
                .Append(" (").Append(HorizontalAlignment).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .verticalAlignment    = ")
                .Append("0x").Append(HexDump.ToHex(VerticalAlignment))
                .Append(" (").Append(VerticalAlignment).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .DisplayMode          = ")
                .Append("0x").Append(HexDump.ToHex(DisplayMode))
                .Append(" (").Append(DisplayMode).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .rgbColor             = ")
                .Append("0x").Append(HexDump.ToHex(RgbColor))
                .Append(" (").Append(RgbColor).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .x                    = ")
                .Append("0x").Append(HexDump.ToHex(X))
                .Append(" (").Append(X).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .y                    = ")
                .Append("0x").Append(HexDump.ToHex(Y))
                .Append(" (").Append(Y).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .width                = ")
                .Append("0x").Append(HexDump.ToHex(Width))
                .Append(" (").Append(Width).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .height               = ")
                .Append("0x").Append(HexDump.ToHex(Height))
                .Append(" (").Append(Height).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options1             = ")
                .Append("0x").Append(HexDump.ToHex(Options1))
                .Append(" (").Append(Options1).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .autoColor                = ").Append(IsAutoColor).Append('\n');
            buffer.Append("         .showKey                  = ").Append(ShowKey).Append('\n');
            buffer.Append("         .showValue                = ").Append(ShowValue).Append('\n');
            //buffer.Append("         .vertical                 = ").Append(IsVertical).Append('\n');
            buffer.Append("         .autoGeneratedText        = ").Append(IsAutoGeneratedText).Append('\n');
            buffer.Append("         .generated                = ").Append(IsGenerated).Append('\n');
            buffer.Append("         .autoLabelDeleted         = ").Append(IsAutoLabelDeleted).Append('\n');
            buffer.Append("         .autoBackground           = ").Append(IsAutoBackground).Append('\n');
            //buffer.Append("         .rotation                 = ").Append(Rotation).Append('\n');
            buffer.Append("         .showCategoryLabelAsPercentage     = ").Append(ShowCategoryLabelAsPercentage).Append('\n');
            buffer.Append("         .showValueAsPercentage     = ").Append(ShowValueAsPercentage).Append('\n');
            buffer.Append("         .showBubbleSizes          = ").Append(ShowBubbleSizes).Append('\n');
            buffer.Append("         .showLabel                = ").Append(ShowLabel).Append('\n');
            buffer.Append("    .IndexOfColorValue    = ")
                .Append("0x").Append(HexDump.ToHex(IndexOfColorValue))
                .Append(" (").Append(IndexOfColorValue).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("    .options2             = ")
                .Append("0x").Append(HexDump.ToHex(Options2))
                .Append(" (").Append(Options2).Append(" )");
            buffer.Append(Environment.NewLine);
            buffer.Append("         .dataLabelPlacement       = ").Append(DataLabelPlacement).Append('\n');
            buffer.Append("         .readingOrder             = ").Append(ReadingOrder).Append('\n');
            buffer.Append("    .textRotation         = ")
                .Append("0x").Append(HexDump.ToHex(TextRotation))
                .Append(" (").Append(TextRotation).Append(" )");
            buffer.Append(Environment.NewLine);

            buffer.Append("[/TEXT]\n");
            return buffer.ToString();
        }

        public override void Serialize(ILittleEndianOutput out1)
        {
            out1.WriteByte(field_1_horizontalAlignment);
            out1.WriteByte(field_2_verticalAlignment);
            out1.WriteShort(field_3_DisplayMode);
            out1.WriteInt(field_4_rgbColor);
            out1.WriteInt(field_5_x);
            out1.WriteInt(field_6_y);
            out1.WriteInt(field_7_width);
            out1.WriteInt(field_8_height);
            out1.WriteShort(field_9_options1);
            out1.WriteShort(field_10_IndexOfColorValue);
            out1.WriteShort(field_11_options2);
            out1.WriteShort(field_12_textRotation);
        }

        /**
         * Size of record (exluding 4 byte header)
         */
        protected override int DataSize
        {
            get { return  1 + 1 + 2 + 4 + 4 + 4 + 4 + 4 + 2 + 2 + 2 + 2; }
        }

        public override short Sid
        {
            get { return sid; }
        }

        public override Object Clone()
        {
            TextRecord rec = new TextRecord();

            rec.field_1_horizontalAlignment = field_1_horizontalAlignment;
            rec.field_2_verticalAlignment = field_2_verticalAlignment;
            rec.field_3_DisplayMode = field_3_DisplayMode;
            rec.field_4_rgbColor = field_4_rgbColor;
            rec.field_5_x = field_5_x;
            rec.field_6_y = field_6_y;
            rec.field_7_width = field_7_width;
            rec.field_8_height = field_8_height;
            rec.field_9_options1 = field_9_options1;
            rec.field_10_IndexOfColorValue = field_10_IndexOfColorValue;
            rec.field_11_options2 = field_11_options2;
            rec.field_12_textRotation = field_12_textRotation;
            return rec;
        }




        /**
         * Get the horizontal alignment field for the Text record.
         *
         * @return  One of 
         *        HORIZONTAL_ALIGNMENT_LEFT
         *        HORIZONTAL_ALIGNMENT_CENTER
         *        HORIZONTAL_ALIGNMENT_BOTTOM
         *        HORIZONTAL_ALIGNMENT_JUSTIFY
         */
        public byte HorizontalAlignment
        {
            get { return field_1_horizontalAlignment; }
            set { this.field_1_horizontalAlignment = value; }
        }

        /**
         * Get the vertical alignment field for the Text record.
         *
         * @return  One of 
         *        VERTICAL_ALIGNMENT_TOP
         *        VERTICAL_ALIGNMENT_CENTER
         *        VERTICAL_ALIGNMENT_BOTTOM
         *        VERTICAL_ALIGNMENT_JUSTIFY
         */
        public byte VerticalAlignment
        {
            get { return field_2_verticalAlignment; }
            set { this.field_2_verticalAlignment = value; }
        }

        /**
         * Get the Display mode field for the Text record.
         *
         * @return  One of 
         *        DISPLAY_MODE_TRANSPARENT
         *        DISPLAY_MODE_OPAQUE
         */
        public short DisplayMode
        {
            get { return field_3_DisplayMode; }
            set { this.field_3_DisplayMode = value; }
        }

        /**
         * Get the rgbColor field for the Text record.
         */
        public int RgbColor
        {
            get { return field_4_rgbColor; }
            set { this.field_4_rgbColor = value; }
        }

        /**
         * Get the x field for the Text record.
         */
        public int X
        {
            get { return field_5_x; }
            set { this.field_5_x = value; }
        }

        /**
         * Get the y field for the Text record.
         */
        public int Y
        {
            get { return field_6_y; }
            set { this.field_6_y = value; }
        }


        /**
         * Set the width field for the Text record.
         */
        public int Width
        {
            get { return field_7_width; }
            set { this.field_7_width = value; }
        }

        /**
         * Get the height field for the Text record.
         */
        public int Height
        {
            get { return field_8_height; }
            set { this.field_8_height = value; }
        }

        /**
         * Get the options1 field for the Text record.
         */
        public short Options1
        {
            get { return field_9_options1; }
            set { this.field_9_options1 = value; }
        }

        /**
         * Get the index of color value field for the Text record.
         */
        public short IndexOfColorValue
        {
            get { return field_10_IndexOfColorValue; }
            set { this.field_10_IndexOfColorValue = value; }
        }

        /**
         * Get the options2 field for the Text record.
         */
        public short Options2
        {
            get { return field_11_options2; }
            set { this.field_11_options2 = value; }
        }

        /**
         * Get the text rotation field for the Text record.
         */
        public short TextRotation
        {
            get { return field_12_textRotation; }
            set { this.field_12_textRotation = value; }
        }

        /**
         * true = automaticly selected colour, false = user-selected
         * @return  the auto color field value.
         */
        public bool IsAutoColor
        {
            get { return autoColor.IsSet(field_9_options1); }
            set { field_9_options1 = autoColor.SetShortBoolean(field_9_options1, value); }
        }

        /**
         * true = draw legend
         * @return  the show key field value.
         */
        public bool ShowKey
        {
            get { return showKey.IsSet(field_9_options1); }
            set { field_9_options1 = showKey.SetShortBoolean(field_9_options1, value); }
        }

        /**
         * false = text is category label
         * @return  the show value field value.
         */
        public bool ShowValue
        {
            get { return showValue.IsSet(field_9_options1); }
            set { field_9_options1 = showValue.SetShortBoolean(field_9_options1, value); }
        }

        /*
         * Sets the vertical field value.
         * true = text is vertical
         */
        /*public bool IsVertical
        {
            get { return vertical.IsSet(field_9_options1); }
            set { field_9_options1 = vertical.SetShortBoolean(field_9_options1, value); }
        }*/

        /**
         * 
         * @return  the auto generated text field value.
         */
        public bool IsAutoGeneratedText
        {
            get { return autoText.IsSet(field_9_options1); }
            set { field_9_options1 = autoText.SetShortBoolean(field_9_options1, value); }
        }

        /**
         * 
         * @return  the generated field value.
         */
        public bool IsGenerated
        {
            get { return generated.IsSet(field_9_options1); }
            set { field_9_options1 = generated.SetShortBoolean(field_9_options1, value); }
        }

        /**
         * 
         * @return  the auto label deleted field value.
         */
        public bool IsAutoLabelDeleted
        {
            get { return autoLabelDeleted.IsSet(field_9_options1); }
            set { field_9_options1 = autoLabelDeleted.SetShortBoolean(field_9_options1, value); }
        }

        /**
         * 
         * @return  the auto background field value.
         */
        public bool IsAutoBackground
        {
            get { return autoBackground.IsSet(field_9_options1); }
            set { field_9_options1 = autoBackground.SetShortBoolean(field_9_options1, value); }
        }

        /*
         * 
         * @return  the rotation field value.
         */
        /*public short Rotation
        {
            get{return rotation.GetShortValue(field_9_options1);}
            set { field_9_options1 = rotation.SetShortValue(field_9_options1, value); }
        }*/

        /**
         * 
         * @return  the show category label as percentage field value.
         */
        public bool ShowCategoryLabelAsPercentage
        {
            get { return showCategoryLabelAsPercentage.IsSet(field_9_options1); }
            set { showCategoryLabelAsPercentage.SetShortBoolean(field_9_options1, value); }
        }

        /**
         * 
         * @return  the show value as percentage field value.
         */
        public bool ShowValueAsPercentage
        {
            get{return showValueAsPercentage.IsSet(field_9_options1);}
            set { field_9_options1 = showValueAsPercentage.SetShortBoolean(field_9_options1, value); }
        }

        /**
         * 
         * @return  the show bubble sizes field value.
         */
        public bool ShowBubbleSizes
        {
            get { return showBubbleSizes.IsSet(field_9_options1); }
            set { field_9_options1 = showBubbleSizes.SetShortBoolean(field_9_options1, value); }
        }
        /**
         * 
         * @return  the show label field value.
         */
        public bool ShowLabel
        {
            get { return showLabel.IsSet(field_9_options1); }
            set { field_9_options1 = showLabel.SetShortBoolean(field_9_options1, value); }
        }

        /**
         * 
         * @return  the data label placement field value.
         */
        public short DataLabelPlacement
        {
            get{return dataLabelPlacement.GetShortValue(field_11_options2);}
            set { field_11_options2 = dataLabelPlacement.SetShortValue(field_11_options2, value); }
        }

        public short ReadingOrder
        {
            get { return readingOrder.GetShortValue(field_11_options2); }
            set { field_11_options2 = readingOrder.SetShortValue(field_11_options2, value); }
        }
    }
}



