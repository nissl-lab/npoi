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

namespace NPOI.SS.UserModel
{
    /// <summary>
    /// High level representation for Font Formatting component
    /// of Conditional Formatting Settings
    /// </summary>
    /// @author Dmitriy Kumshayev
    /// @author Yegor Kozlov
    public interface IFontFormatting
    {
        /// <summary>
        /// get or set the type of super or subscript for the font
        /// </summary>
        FontSuperScript EscapementType
        {
            get;
            set;
        }
        /// <summary>
        /// get or set font color index
        /// </summary>
        short FontColorIndex
        {
            get;
            set;
        }
        /// <summary>
        /// get or set the colour of the font, or null if no colour applied
        /// </summary>
        IColor FontColor
        {
            get;
            set;
        }
        /// <summary>
        /// get or set the height of the font in 1/20th point units
        /// </summary>
        int FontHeight
        {
            get;
            set;
        }
        /// <summary>
        /// get or set the type of underlining for the font
        /// </summary>
        FontUnderlineType UnderlineType
        {
            get;
            set;
        }


        /// <summary>
        /// Get whether the font weight is Set to bold or not
        /// </summary>
        /// <returns>bold - whether the font is bold or not</returns>
        bool IsBold { get; }

        /// <summary>
        /// </summary>
        /// <returns>true if font style was Set to <i>italic</i></returns>
        bool IsItalic { get; }

        /// <summary>
        /// Set font style options.
        /// </summary>
        /// <param name="italic">- if true, Set posture style to italic, otherwise to normal</param>
        /// <param name="bold">if true, Set font weight to bold, otherwise to normal</param>
        void SetFontStyle(bool italic, bool bold);

        /// <summary>
        /// Set font style options to default values (non-italic, non-bold)
        /// </summary>
        void ResetFontStyle();
    }

}