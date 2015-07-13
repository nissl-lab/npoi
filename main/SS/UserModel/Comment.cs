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
    public interface IComment
    {
        /**
         * Sets whether this comment is visible.
         *
         * @return <c>true</c> if the comment is visible, <c>false</c> otherwise
         */
        bool Visible { get; set; }

        /**
         * Return the row of the cell that Contains the comment
         *
         * @return the 0-based row of the cell that Contains the comment
         */
        int Row { get; set; }


        /**
         * Return the column of the cell that Contains the comment
         *
         * @return the 0-based column of the cell that Contains the comment
         */
        int Column { get; set; }


        /**
         * Name of the original comment author
         *
         * @return the name of the original author of the comment
         */
        String Author { get; set; }

        /**
         * Fetches the rich text string of the comment
         */
        IRichTextString String { get; set; }

        /**
         * Return defines position of this anchor in the sheet.
         *
         * @return defines position of this anchor in the sheet
         */
        IClientAnchor ClientAnchor { get; }
    }
}
