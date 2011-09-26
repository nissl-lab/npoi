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
    using NPOI.SS.UserModel;
    /// <summary>
    /// An abstract shape.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    [Serializable]
    public abstract class HSSFShape:IShape
    {
        public static int LINEWIDTH_ONE_PT = 12700;
        public static int LINEWIDTH_DEFAULT = 9525;

        public static int LINESTYLE_SOLID = 0;              // Solid (continuous) pen
        public static int LINESTYLE_DASHSYS = 1;            // PS_DASH system   dash style
        public static int LINESTYLE_DOTSYS = 2;             // PS_DOT system   dash style
        public static int LINESTYLE_DASHDOTSYS = 3;         // PS_DASHDOT system dash style
        public static int LINESTYLE_DASHDOTDOTSYS = 4;      // PS_DASHDOTDOT system dash style
        public static int LINESTYLE_DOTGEL = 5;             // square dot style
        public static int LINESTYLE_DASHGEL = 6;            // dash style
        public static int LINESTYLE_LONGDASHGEL = 7;        // long dash style
        public static int LINESTYLE_DASHDOTGEL = 8;         // dash short dash
        public static int LINESTYLE_LONGDASHDOTGEL = 9;     // long dash short dash
        public static int LINESTYLE_LONGDASHDOTDOTGEL = 10; // long dash short dash short dash
        public static int LINESTYLE_NONE = -1;

        IShape parent;
        HSSFAnchor anchor;
        int lineStyleColor = 0x08000040;
        int fillColor = 0x08000009;
        int lineWidth = LINEWIDTH_DEFAULT;    // 12700 = 1pt
        int lineStyle = LINESTYLE_SOLID;
        bool noFill = false;

        /// <summary>
        /// Create a new shape with the specified parent and anchor.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <param name="anchor">The anchor.</param>
        protected HSSFShape(IShape parent, HSSFAnchor anchor)
        {
            this.parent = parent;
            this.Anchor = anchor;
        }

        /// <summary>
        /// Gets the parent shape.
        /// </summary>
        /// <value>The parent.</value>
        public IShape Parent
        {
            get { return parent; }
        }

        /// <summary>
        /// Gets or sets the anchor that is used by this shape.
        /// </summary>
        /// <value>The anchor.</value>
        public HSSFAnchor Anchor
        {
            get { return anchor; }
            set
            {
                if (parent == null)
                {
                    if (value is HSSFChildAnchor)
                        throw new ArgumentException("Must use client anchors for shapes directly attached to sheet.");
                }
                else
                {
                    if (value is HSSFClientAnchor)
                        throw new ArgumentException("Must use child anchors for shapes attached to Groups.");
                }

                this.anchor = value;
            }
        }

        public void SetLineStyleColor(int lineStyleColor)
        {
            this.lineStyleColor = lineStyleColor;
        }
        /// <summary>
        /// The color applied to the lines of this shape.
        /// </summary>
        /// <value>The color of the line style.</value>
        public int LineStyleColor
        {
            get
            {
                return lineStyleColor;
            }
        }

        /// <summary>
        /// Sets the color applied to the lines of this shape
        /// </summary>
        /// <param name="red">The red.</param>
        /// <param name="green">The green.</param>
        /// <param name="blue">The blue.</param>
        public void SetLineStyleColor(int red, int green, int blue)
        {
            this.lineStyleColor = ((blue) << 16) | ((green) << 8) | red;
        }

        /// <summary>
        /// Gets or sets the color used to fill this shape.
        /// </summary>
        /// <value>The color of the fill.</value>
        public int FillColor
        {
            get{return fillColor;}
            set { fillColor = value; }
        }

        /// <summary>
        /// Sets the color used to fill this shape.
        /// </summary>
        /// <param name="red">The red.</param>
        /// <param name="green">The green.</param>
        /// <param name="blue">The blue.</param>
        public void SetFillColor(int red, int green, int blue)
        {
            this.FillColor = ((blue) << 16) | ((green) << 8) | red;
        }

        /// <summary>
        /// Gets or sets with width of the line in EMUs.  12700 = 1 pt.
        /// </summary>
        /// <value>The width of the line.</value>
        public int LineWidth
        {
            get { return lineWidth; }
            set { this.lineWidth = value; }
        }

        /// <summary>
        /// Gets or sets One of the constants in LINESTYLE_*
        /// </summary>
        /// <value>The line style.</value>
        public int LineStyle
        {
            get { return lineStyle; }
            set { lineStyle = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is no fill.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this shape Is not filled with a color; otherwise, <c>false</c>.
        /// </value>
        public bool IsNoFill
        {
            get { return noFill; }
            set { this.noFill = value; }
        }

        /// <summary>
        /// Count of all children and their childrens children.
        /// </summary>
        /// <value>The count of all children.</value>
        public virtual int CountOfAllChildren
        {
            get { return 1; }
        }
    }
}