/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

namespace NPOI.SS.UserModel
{
    public enum BorderFormattingStyle:short
    {
        /** No border */
        BORDER_NONE = 0x0,
        /** Thin border */
        BORDER_THIN = 0x1,
        /** Medium border */
        BORDER_MEDIUM = 0x2,
        /** dash border */
        BORDER_DASHED = 0x3,
        /** dot border */
        BORDER_HAIR = 0x4,
        /** Thick border */
        BORDER_THICK = 0x5,
        /** double-line border */
        BORDER_DOUBLE = 0x6,
        /** hair-line border */
        BORDER_DOTTED = 0x7,
        /** Medium dashed border */
        BORDER_MEDIUM_DASHED = 0x8,
        /** dash-dot border */
        BORDER_DASH_DOT = 0x9,
        /** medium dash-dot border */
        BORDER_MEDIUM_DASH_DOT = 0xA,
        /** dash-dot-dot border */
        BORDER_DASH_DOT_DOT = 0xB,
        /** medium dash-dot-dot border */
        BORDER_MEDIUM_DASH_DOT_DOT = 0xC,
        /** slanted dash-dot border */
        BORDER_SLANTED_DASH_DOT = 0xD
    }
    /**
     * @author Dmitriy Kumshayev
     * @author Yegor Kozlov
     */
    public interface IBorderFormatting
    {
        short BorderBottom { get; set; }

        short BorderDiagonal { get; set; }

        short BorderLeft { get; set; }

        short BorderRight { get; set; }

        short BorderTop { get; set; }

        short BottomBorderColor { get; set; }

        short DiagonalBorderColor { get; set; }

        short LeftBorderColor { get; set; }

        short RightBorderColor { get; set; }

        short TopBorderColor { get; set; }
    }

}