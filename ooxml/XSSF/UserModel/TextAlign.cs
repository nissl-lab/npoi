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
     * Specified a list of text alignment types
     */
    public enum TextAlign
    {
        None,
        /**
         * Align text to the left margin.
         */
        LEFT,
        /**
         * Align text in the center.
         */
        CENTER,

        /**
         * Align text to the right margin.
         */
        RIGHT,

        /**
         * Align text so that it is justified across the whole line. It
         * is smart in the sense that it will not justify sentences
         * which are short
         */
        JUSTIFY,
        JUSTIFY_LOW,
        DIST,
        THAI_DIST
    }

}