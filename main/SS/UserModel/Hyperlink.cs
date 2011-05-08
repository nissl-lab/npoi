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
    using System;

    public enum HyperlinkType : int
    {
        Unknown = 0,
        /// <summary>
        /// Link to a existing file or web page
        /// </summary>
        URL = 1,
        /// <summary>
        /// Link to a place in this document
        /// </summary>
        DOCUMENT = 2,
        /// <summary>
        /// Link to an E-mail Address
        /// </summary>
        EMAIL = 3,
        /// <summary>
        /// Link to a file
        /// </summary>
        FILE = 4
    }
    /**
     * Represents an Excel hyperlink.
     */
    public interface Hyperlink  //NPOI.COMMON.UserModel.Hyperlink
    {
        /**
 * Hypelink address. Depending on the hyperlink type it can be URL, e-mail, patrh to a file, etc.
 *
 * @return  the address of this hyperlink
 */
        String Address { get; set; }

        /**
         * Return text label for this hyperlink
         *
         * @return  text to display
         */
        String Label { get; set; }

        /**
         * Return the type of this hyperlink
         *
         * @return the type of this hyperlink
         */
        HyperlinkType Type { get; }
        /**
         * Return the row of the first cell that Contains the hyperlink
         *
         * @return the 0-based row of the cell that Contains the hyperlink
         */
        int FirstRow { get; set; }
        /**
         * Return the row of the last cell that Contains the hyperlink
         *
         * @return the 0-based row of the last cell that Contains the hyperlink
         */
        int LastRow { get; set; }

        /**
         * Return the column of the first cell that Contains the hyperlink
         *
         * @return the 0-based column of the first cell that Contains the hyperlink
         */
        int FirstColumn { get; set; }

        /**
         * Return the column of the last cell that Contains the hyperlink
         *
         * @return the 0-based column of the last cell that Contains the hyperlink
         */
        int LastColumn { get; set; }

        string TextMark { get; set; }
    }

}