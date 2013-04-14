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

namespace NPOI.SS.UserModel
{

    /**
     * The enumeration value indicating the style of fill pattern being used for a cell format.
     * 
     */
    public enum FillPattern : short
    {

        /**  No background */
        NoFill = 0,

        /**  Solidly Filled */
        SolidForeground = 1,

        /**  Small fine dots */
        FineDots = 2,

        /**  Wide dots */
        AltBars = 3,

        /**  Sparse dots */
        SparseDots = 4,

        /**  Thick horizontal bands */
        ThickHorizontalBands = 5,

        /**  Thick vertical bands */
        ThickVerticalBands = 6,

        /**  Thick backward facing diagonals */
        ThickBackwardDiagonals = 7,

        /**  Thick forward facing diagonals */
        ThickForwardDiagonals = 8,

        /**  Large spots */
        BigSpots = 9,

        /**  Brick-like layout */
        Bricks = 10,

        /**  Thin horizontal bands */
        ThinHorizontalBands = 11,

        /**  Thin vertical bands */
        ThinVerticalBands = 12,

        /**  Thin backward diagonal */
        ThinBackwardDiagonals = 13,

        /**  Thin forward diagonal */
        ThinForwardDiagonals = 14,

        /**  Squares */
        Squares = 15,

        /**  Diamonds */
        Diamonds = 16,

        /**  Less Dots */
        LessDots = 17,

        /**  Least Dots */
        LeastDots = 18
    }
}
