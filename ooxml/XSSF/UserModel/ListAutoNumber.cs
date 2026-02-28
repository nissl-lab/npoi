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
namespace NPOI.XSSF.UserModel
{
    /**
     * Specifies type of automatic numbered bullet points that should be applied to a paragraph.
     */
    public enum ListAutoNumber
    {
        /**
         * (a), (b), (c), ...
         */
        ALPHA_LC_PARENT_BOTH,
        /**
         * (A), (B), (C), ...
         */
        ALPHA_UC_PARENT_BOTH,
        /**
         * a), b), c), ...
         */
        ALPHA_LC_PARENT_R,
        /**
         * A), B), C), ...
         */
        ALPHA_UC_PARENT_R,
        /**
         *  a., b., c., ...
         */
        ALPHA_LC_PERIOD,
        /**
         * A., B., C., ...
         */
        ALPHA_UC_PERIOD,
        /**
         * (1), (2), (3), ...
         */
        ARABIC_PARENT_BOTH,
        /**
         * 1), 2), 3), ...
         */
        ARABIC_PARENT_R,

        /**
         * 1., 2., 3., ...
         */
        ARABIC_PERIOD,
        /**
         * 1, 2, 3, ...
         */
        ARABIC_PLAIN,

        /**
         * (i), (ii), (iii), ...
         */
        ROMAN_LC_PARENT_BOTH,
        /**
         * (I), (II), (III), ...
         */
        ROMAN_UC_PARENT_BOTH,
        /**
         * i), ii), iii), ...
         */
        ROMAN_LC_PARENT_R,
        /**
         * I), II), III), ...
         */
        ROMAN_UC_PARENT_R,
        /**
         *  i., ii., iii., ...
         */
        ROMAN_LC_PERIOD,
        /**
         * I., II., III., ...
         */
        ROMAN_UC_PERIOD,
        /**
         * Dbl-byte circle numbers
         */
        CIRCLE_NUM_DB_PLAIN,
        /**
         * Wingdings black circle numbers
         */
        CIRCLE_NUM_WD_BLACK_PLAIN,
        /**
         * Wingdings white circle numbers
         */
        CIRCLE_NUM_WD_WHITE_PLAIN
    }
}