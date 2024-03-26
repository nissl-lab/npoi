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
    /// <summary>
    /// Rich text unicode string.  These strings can have fonts
    ///  applied to arbitary parts of the string.
    /// </summary>
    /// @author Glen Stampoultzis (glens at apache.org)
    /// @author Jason Height (jheight at apache.org)
    public interface IRichTextString
    {

        /// <summary>
        /// Applies a font to the specified characters of a string.
        /// </summary>
        /// <param name="startIndex">   The start index to apply the font to (inclusive)</param>
        /// <param name="endIndex">     The end index to apply the font to (exclusive)</param>
        /// <param name="fontIndex">    The font to use.</param>
        void ApplyFont(int startIndex, int endIndex, short fontIndex);

        /// <summary>
        /// Applies a font to the specified characters of a string.
        /// </summary>
        /// <param name="startIndex">   The start index to apply the font to (inclusive)</param>
        /// <param name="endIndex">     The end index to apply to font to (exclusive)</param>
        /// <param name="font">         The index of the font to use.</param>
        void ApplyFont(int startIndex, int endIndex, IFont font);

        /// <summary>
        /// Sets the font of the entire string.
        /// </summary>
        /// <param name="font">         The font to use.</param>
        void ApplyFont(IFont font);

        /// <summary>
        /// Removes any formatting that may have been applied to the string.
        /// </summary>
        void ClearFormatting();

        /// <summary>
        /// Returns the plain string representation.
        /// </summary>
        String String { get; }

        /// <summary>
        /// </summary>
        /// <returns>the number of characters in the font.</returns>
        int Length { get; }

        /// <summary>
        /// </summary>
        /// <returns>The number of formatting Runs used.</returns>
        ///
        int NumFormattingRuns { get; }

        /// <summary>
        /// The index within the string to which the specified formatting run applies.
        /// </summary>
        /// <param name="index">    the index of the formatting run</param>
        /// <returns>the index within the string.</returns>
        int GetIndexOfFormattingRun(int index);

        /// <summary>
        /// Applies the specified font to the entire string.
        /// </summary>
        /// <param name="fontIndex"> the font to apply.</param>
        void ApplyFont(short fontIndex);
    }
}