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
        /// Link to an existing file or web page
        /// </summary>
        Url = 1,
        /// <summary>
        /// Link to a place in this document
        /// </summary>
        Document = 2,
        /// <summary>
        /// Link to an E-mail Address
        /// </summary>
        Email = 3,
        /// <summary>
        /// Link to a file
        /// </summary>
        File = 4
    }
    /// <summary>
    /// Represents an Excel hyperlink.
    /// </summary>
    public interface IHyperlink
    {
        /// <summary>
        /// Hyperlink address. Depending on the hyperlink type it can be URL, e-mail, patrh to a file, etc.
        /// </summary>
        String Address { get; set; }

        /// <summary>
        /// text label for this hyperlink
        /// </summary>
        String Label { get; set; }

        /// <summary>
        /// the type of this hyperlink
        /// </summary>
        HyperlinkType Type { get; }

        /// <summary>
        /// the row of the first cell that Contains the hyperlink
        /// </summary>
        int FirstRow { get; set; }
        /// <summary>
        /// the row of the last cell that Contains the hyperlink
        /// </summary>
        int LastRow { get; set; }

        /// <summary>
        /// the column of the first cell that Contains the hyperlink
        /// </summary>
        int FirstColumn { get; set; }

        /// <summary>
        /// the column of the last cell that Contains the hyperlink
        /// </summary>
        int LastColumn { get; set; }

        string TextMark { get; set; }
    }

}