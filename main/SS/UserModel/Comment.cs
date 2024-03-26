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
    using NPOI.SS.Util;
    using System;
    public interface IComment
    {
        /// <summary>
        /// Sets whether this comment is visible.
        /// </summary>
        /// <returns><c>true</c> if the comment is visible, <c>false</c> otherwise</returns>
        bool Visible { get; set; }

        /// <summary>
        /// Get or set the address of the cell that this comment is attached to
        /// </summary>
        CellAddress Address { get; set; }

        /// <summary>
        /// Set the address of the cell that this comment is attached to
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">col</param>
        void SetAddress(int row, int col);

        /// <summary>
        /// Return the row of the cell that Contains the comment
        /// </summary>
        /// <returns>the 0-based row of the cell that Contains the comment</returns>
        int Row { get; set; }


        /// <summary>
        /// Return the column of the cell that Contains the comment
        /// </summary>
        /// <returns>the 0-based column of the cell that Contains the comment</returns>
        int Column { get; set; }


        /// <summary>
        /// Name of the original comment author
        /// </summary>
        /// <returns>the name of the original author of the comment</returns>
        String Author { get; set; }

        /// <summary>
        /// Fetches the rich text string of the comment
        /// </summary>
        IRichTextString String { get; set; }

        /// <summary>
        /// <para>
        /// Return defines position of this anchor in the sheet.
        /// </para>
        /// <para>
        /// The anchor is the yellow box/balloon that is rendered on top of the sheets
        /// when the comment is visible.
        /// </para>
        /// <para>
        /// To associate a comment with a different cell, use <see cref="Address"/>.
        /// </para>
        /// </summary>
        /// <returns>defines position of this anchor in the sheet</returns>
        IClientAnchor ClientAnchor { get; }
    }
}
