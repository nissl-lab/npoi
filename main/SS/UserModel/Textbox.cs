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

using System;

namespace NPOI.SS.UserModel
{
    /// <summary>
    /// This interface isn't implemented ...
    /// </summary>
    [Obsolete("deprecated in POI 3.16 beta1, scheduled for removal in 3.18")]
    public interface ITextbox : IShape
    {

        //public const short OBJECT_TYPE_TEXT = 6;

        /// <summary>
        /// </summary>
        /// <returns>the rich text string for this textbox.</returns>
        IRichTextString String { get; set; }


        /// <summary>
        /// </summary>
        /// <returns>Returns the left margin within the textbox.</returns>
        int MarginLeft { get; set; }


        /// <summary>
        /// </summary>
        /// <returns>returns the right margin within the textbox.</returns>
        int MarginRight { get; set; }


        /// <summary>
        /// </summary>
        /// <returns>returns the top margin within the textbox.</returns>
        int MarginTop { get; set; }

        /// <summary>
        /// s the bottom margin within the textbox.
        /// </summary>
        int MarginBottom { get; set; }

        short HorizontalAlignment { get; set; }
        short VerticalAlignment { get; set; }
    }
}
