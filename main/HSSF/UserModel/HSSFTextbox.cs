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
    using NPOI.Util;


    /// <summary>
    /// A textbox Is a shape that may hold a rich text string.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    [Serializable]
    public class HSSFTextbox : HSSFSimpleShape //,NPOI.SS.UserModel.ITextbox
    {
        public static short OBJECT_TYPE_TEXT = 6;

            /**
     * How to align text horizontally
     */
    public static short  HORIZONTAL_ALIGNMENT_LEFT = 1;
    public static short  HORIZONTAL_ALIGNMENT_CENTERED = 2;
    public static short  HORIZONTAL_ALIGNMENT_RIGHT = 3;
    public static short  HORIZONTAL_ALIGNMENT_JUSTIFIED = 4;
    public static short  HORIZONTAL_ALIGNMENT_DISTRIBUTED = 7;

    /**
     * How to align text vertically
     */
    public static short  VERTICAL_ALIGNMENT_TOP    = 1;
    public static short  VERTICAL_ALIGNMENT_CENTER = 2;
    public static short  VERTICAL_ALIGNMENT_BOTTOM = 3;
    public static short  VERTICAL_ALIGNMENT_JUSTIFY = 4;
    public static short  VERTICAL_ALIGNMENT_DISTRIBUTED= 7;

        int marginLeft, marginRight, marginTop, marginBottom;
        short halign;
        short valign;

        NPOI.SS.UserModel.IRichTextString str = new HSSFRichTextString("");

        /// <summary>
        /// Construct a new textbox with the given parent and anchor.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <param name="anchor">One of HSSFClientAnchor or HSSFChildAnchor</param>
        public HSSFTextbox(HSSFShape parent, HSSFAnchor anchor):base(parent, anchor)
        {
            
            ShapeType = (OBJECT_TYPE_TEXT);

            halign = HORIZONTAL_ALIGNMENT_LEFT;
            valign = VERTICAL_ALIGNMENT_TOP;
        }

        /// <summary>
        /// Gets or sets the rich text string for this textbox.
        /// </summary>
        /// <value>The string.</value>
        public virtual NPOI.SS.UserModel.IRichTextString String
        {
            get { return str; }
            set
            {             //if font Is not Set we must Set the default one
                if (value.NumFormattingRuns == 0) value.ApplyFont((short)0);

                this.str = value;
            }
        }
        /// <summary>
        /// Gets or sets the left margin within the textbox.
        /// </summary>
        /// <value>The margin left.</value>
        public int MarginLeft
        {
            get { return marginLeft; }
            set { this.marginLeft = value; }
        }


        /// <summary>
        /// Gets or sets the right margin within the textbox.
        /// </summary>
        /// <value>The margin right.</value>
        public int MarginRight
        {
            get { return marginRight; }
            set { this.marginRight = value; }
        }

        /// <summary>
        /// Gets or sets the top margin within the textbox
        /// </summary>
        /// <value>The top margin.</value>
        public int MarginTop
        {
            get { return marginTop; }
            set { this.marginTop = value; }
        }

        /// <summary>
        /// Gets or sets the bottom margin within the textbox.
        /// </summary>
        /// <value>The margin bottom.</value>
        public int MarginBottom
        {
            get { return marginBottom; }
            set { this.marginBottom = value; }
        }

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        /// <value>The horizontal alignment.</value>
        public short HorizontalAlignment
        {
            get { return halign; }
            set { this.halign = value; }
        }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        /// <value>The vertical alignment.</value>
        public short VerticalAlignment
        {
            get { return valign; }
            set { this.valign = value; }
        }
    }
}