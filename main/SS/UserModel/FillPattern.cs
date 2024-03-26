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

    /// <summary>
    /// The enumeration value indicating the style of fill pattern being used for a cell format.
    /// </summary>
    public enum FillPattern : short
    {

        /// <summary>
        ///  No background */
        /// </summary>
        NoFill = 0,

        /// <summary>
        ///  Solidly Filled */
        /// </summary>
        SolidForeground = 1,

        /// <summary>
        ///  Small fine dots */
        /// </summary>
        FineDots = 2,

        /// <summary>
        ///  Wide dots */
        /// </summary>
        AltBars = 3,

        /// <summary>
        ///  Sparse dots */
        /// </summary>
        SparseDots = 4,

        /// <summary>
        ///  Thick horizontal bands */
        /// </summary>
        ThickHorizontalBands = 5,

        /// <summary>
        ///  Thick vertical bands */
        /// </summary>
        ThickVerticalBands = 6,

        /// <summary>
        ///  Thick backward facing diagonals */
        /// </summary>
        ThickBackwardDiagonals = 7,

        /// <summary>
        ///  Thick forward facing diagonals */
        /// </summary>
        ThickForwardDiagonals = 8,

        /// <summary>
        ///  Large spots */
        /// </summary>
        BigSpots = 9,

        /// <summary>
        ///  Brick-like layout */
        /// </summary>
        Bricks = 10,

        /// <summary>
        ///  Thin horizontal bands */
        /// </summary>
        ThinHorizontalBands = 11,

        /// <summary>
        ///  Thin vertical bands */
        /// </summary>
        ThinVerticalBands = 12,

        /// <summary>
        ///  Thin backward diagonal */
        /// </summary>
        ThinBackwardDiagonals = 13,

        /// <summary>
        ///  Thin forward diagonal */
        /// </summary>
        ThinForwardDiagonals = 14,

        /// <summary>
        ///  Squares */
        /// </summary>
        Squares = 15,

        /// <summary>
        ///  Diamonds */
        /// </summary>
        Diamonds = 16,

        /// <summary>
        ///  Less Dots */
        /// </summary>
        LessDots = 17,

        /// <summary>
        ///  Least Dots */
        /// </summary>
        LeastDots = 18
    }
}
