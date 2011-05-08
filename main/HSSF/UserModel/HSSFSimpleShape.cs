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
    /// <summary>
    /// Represents a simple shape such as a line, rectangle or oval.
    /// @author Glen Stampoultzis (glens at apache.org)
    /// </summary>
    [Serializable]
    public class HSSFSimpleShape
       : HSSFShape
    {
        // The commented out ones haven't been tested yet or aren't supported
        // by HSSFSimpleShape.

        public const short OBJECT_TYPE_LINE = 1;
        public const short OBJECT_TYPE_RECTANGLE = 2;
        public const short OBJECT_TYPE_OVAL = 3;
        //    public static short       OBJECT_TYPE_ARC                = 4;
        //    public static short       OBJECT_TYPE_CHART              = 5;
        //    public static short       OBJECT_TYPE_TEXT               = 6;
        //    public static short       OBJECT_TYPE_BUTTON             = 7;
        public const short OBJECT_TYPE_PICTURE = 8;
        //    public static short       OBJECT_TYPE_POLYGON            = 9;
        //    public static short       OBJECT_TYPE_CHECKBOX           = 11;
        //    public static short       OBJECT_TYPE_OPTION_BUTTON      = 12;
        //    public static short       OBJECT_TYPE_EDIT_BOX           = 13;
        //    public static short       OBJECT_TYPE_LABEL              = 14;
        //    public static short       OBJECT_TYPE_DIALOG_BOX         = 15;
        //    public static short       OBJECT_TYPE_SPINNER            = 16;
        //    public static short       OBJECT_TYPE_SCROLL_BAR         = 17;
        //    public static short       OBJECT_TYPE_LIST_BOX           = 18;
        //    public static short       OBJECT_TYPE_GROUP_BOX          = 19;
        public const short OBJECT_TYPE_COMBO_BOX = 20;
        public const short OBJECT_TYPE_COMMENT = 25;
        //    public static short       OBJECT_TYPE_MICROSOFT_OFFICE_DRAWING = 30;

        int shapeType = OBJECT_TYPE_LINE;

        /// <summary>
        /// Initializes a new instance of the <see cref="HSSFSimpleShape"/> class.
        /// </summary>
        /// <param name="parent">The parent.</param>
        /// <param name="anchor">The anchor.</param>
        public HSSFSimpleShape(HSSFShape parent, HSSFAnchor anchor):base(parent, anchor)
        {
            
        }

        /// <summary>
        /// Gets the shape type.
        /// </summary>
        /// <value>One of the OBJECT_TYPE_* constants.</value>
        /// @see #OBJECT_TYPE_LINE
        /// @see #OBJECT_TYPE_OVAL
        /// @see #OBJECT_TYPE_RECTANGLE
        /// @see #OBJECT_TYPE_PICTURE
        /// @see #OBJECT_TYPE_COMMENT
        public int ShapeType 
        {
            get
            {
                return shapeType;
            }
            set 
            {
                this.shapeType = value;
            }
        }

    }
}