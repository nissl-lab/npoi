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
    public enum FillPatternType
    {

        /**  No background */
        NO_FILL,

        /**  Solidly Filled */
        SOLID_FOREGROUND,

        /**  Small fine dots */
        FINE_DOTS,

        /**  Wide dots */
        ALT_BARS,

        /**  Sparse dots */
        SPARSE_DOTS,

        /**  Thick horizontal bands */
        THICK_HORZ_BANDS,

        /**  Thick vertical bands */
        THICK_VERT_BANDS,

        /**  Thick backward facing diagonals */
        THICK_BACKWARD_DIAG,

        /**  Thick forward facing diagonals */
        THICK_FORWARD_DIAG,

        /**  Large spots */
        BIG_SPOTS,

        /**  Brick-like layout */
        BRICKS,

        /**  Thin horizontal bands */
        THIN_HORZ_BANDS,

        /**  Thin vertical bands */
        THIN_VERT_BANDS,

        /**  Thin backward diagonal */
        THIN_BACKWARD_DIAG,

        /**  Thin forward diagonal */
        THIN_FORWARD_DIAG,

        /**  Squares */
        SQUARES,

        /**  Diamonds */
        DIAMONDS,

        /**  Less Dots */
        LESS_DOTS,

        /**  Least Dots */
        LEAST_DOTS

    }
}
