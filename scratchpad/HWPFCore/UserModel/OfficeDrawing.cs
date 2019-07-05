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
namespace NPOI.HWPF.UserModel
{


    public enum HorizontalPositioning
    {

        /**
         * The shape is horizontally offset by an absolute distance from the
         * page element.
         */
        ABSOLUTE,

        /**
         * The shape is horizontally positioned at the center of the page
         * element.
         */
        CENTER,

        /**
         * The shape is horizontally positioned like {@link #LEFT} on
         * odd-numbered pages and like {@link #RIGHT} on even-numbered pages.
         */
        INSIDE,

        /**
         * The shape is horizontally positioned at the left side of the page
         * element.
         */
        LEFT,

        /**
         * The shape is horizontally positioned like {@link #RIGHT} on
         * odd-numbered pages and like {@link #LEFT} on even-numbered pages.
         */
        OUTSIDE,

        /**
         * The shape is horizontally positioned at the right side of the page
         * element.
         */
        RIGHT
    }

    public enum HorizontalRelativeElement
    {
        CHAR, MARGIN, PAGE, TEXT
    }

    public enum VerticalPositioning
    {

        /**
         * The shape is vertically offset by an absolute distance from the page
         * element
         */
        ABSOLUTE,

        /**
         * The shape is vertically positioned at the bottom of the page element
         */
        BOTTOM,

        /**
         * The shape is vertically positioned in the center of the page element
         */
        CENTER,

        /**
         * The shape is vertically positioned like msopvTop on odd-numbered
         * pages and like msopvBottom on even-numbered pages
         */
        INSIDE,

        /**
         * The shape is vertically positioned like {@link #BOTTOM} on
         * odd-numbered pages and like {@link #TOP} on even-numbered pages
         */
        OUTSIDE,

        /**
         * The shape is vertically positioned at the top of the page element
         */
        TOP
    }

    public enum VerticalRelativeElement
    {
        LINE, MARGIN, PAGE, TEXT
    }
    /**
     * User-friendly interface to office drawing objects.
     * <p>
     * Some properties and enumeration constants description are quotes from the
     * following sources:
     * <ul>
     * <li>[MS-ODRAW] -- v20110608; Office Drawing Binary File Format; Copyright (c)
     * 2011 Microsoft Corporation.
     * </ul>
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    public interface OfficeDrawing
    {



        /**
         * Returns the type of horizontal positioning to use for a shape
         * 
         * @return the type of horizontal positioning to use for a shape
         */
        HorizontalPositioning GetHorizontalPositioning();

        /**
         * Specifies a page element relative to which a shape is horizontally
         * positioned
         * 
         * @return a page element relative to which a shape is horizontally
         *         positioned
         */
        HorizontalRelativeElement GetHorizontalRelative();

        /**
         * Returns picture data if this shape has (Single?) associated picture data
         */
        byte[] GetPictureData();

        /**
         * Bottom of the rectangle enclosing shape relative to the origin of the
         * shape
         */
        int GetRectangleBottom();

        /**
         * Left of rectangle enclosing shape relative to the origin of the shape
         */
        int GetRectangleLeft();

        /**
         * Right of rectangle enclosing shape relative to the origin of the shape
         */
        int GetRectangleRight();

        /**
         * Top of rectangle enclosing shape relative to the origin of the shape
         */
        int GetRectangleTop();

        /**
         * Shape Identifier
         */
        int GetShapeId();

        /**
         * Specifies the type of vertical positioning to use for a shape
         * 
         * @return return the type of vertical positioning to use for a shape
         */
        VerticalPositioning GetVerticalPositioning();

        /**
         * Specifies a page element relative to which a shape is vertically
         * positioned
         * 
         * @return a page element relative to which a shape is vertically positioned
         */
        VerticalRelativeElement GetVerticalRelativeElement();

    }


}