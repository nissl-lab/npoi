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
    /// <summary>
    /// Used by HSSFPrintSetup.CellError property
    /// </summary>
    public enum DisplayCellErrorType : short
    {
        ErrorDisplayOnSheet = 0,
        ErrorAsBlank = 1,
        ErrorAsDashes = 2,
        ErrorAsNA = 3
    }



    public interface IPrintSetup
    {

        /// <summary>
        /// Returns the paper size.
        /// </summary>
        /// <returns>paper size</returns>
        short PaperSize { get; set; }

        /// <summary>
        /// Returns the scale.
        /// </summary>
        /// <returns>scale</returns>
        short Scale { get; set; }

        /// <summary>
        /// Returns the page start.
        /// </summary>
        /// <returns>page start</returns>
        short PageStart { get; set; }

        /// <summary>
        /// Returns the number of pages wide to fit sheet in.
        /// </summary>
        /// <returns>number of pages wide to fit sheet in</returns>
        short FitWidth { get; set; }

        /// <summary>
        /// Returns the number of pages high to fit the sheet in.
        /// </summary>
        /// <returns>number of pages high to fit the sheet in</returns>
        short FitHeight { get; set; }

        /// <summary>
        /// Returns the left to right print order.
        /// </summary>
        /// <returns>left to right print order</returns>
        bool LeftToRight { get; set; }

        /// <summary>
        /// Returns the landscape mode.
        /// </summary>
        /// <returns>landscape mode</returns>
        bool Landscape { get; set; }

        /// <summary>
        /// Returns the valid Settings.
        /// </summary>
        /// <returns>valid Settings</returns>
        bool ValidSettings { get; set; }

        /// <summary>
        /// Returns the black and white Setting.
        /// </summary>
        /// <returns>black and white Setting</returns>
        bool NoColor { get; set; }

        /// <summary>
        /// Returns the draft mode.
        /// </summary>
        /// <returns>draft mode</returns>
        bool Draft { get; set; }

        /// <summary>
        /// Returns the print notes.
        /// </summary>
        /// <returns>print notes</returns>
        bool Notes { get; set; }

        /// <summary>
        /// Returns the no orientation.
        /// </summary>
        /// <returns>no orientation</returns>
        bool NoOrientation { get; set; }

        /// <summary>
        /// Returns the use page numbers.
        /// </summary>
        /// <returns>use page numbers</returns>
        bool UsePage { get; set; }

        /// <summary>
        /// Returns the horizontal resolution.
        /// </summary>
        /// <returns>horizontal resolution</returns>
        short HResolution { get; set; }

        /// <summary>
        /// Returns the vertical resolution.
        /// </summary>
        /// <returns>vertical resolution</returns>
        short VResolution { get; set; }

        /// <summary>
        /// Returns the header margin.
        /// </summary>
        /// <returns>header margin</returns>
        double HeaderMargin { get; set; }

        /// <summary>
        /// Returns the footer margin.
        /// </summary>
        /// <returns>footer margin</returns>
        double FooterMargin { get; set; }

        /// <summary>
        /// Returns the number of copies.
        /// </summary>
        /// <returns>number of copies</returns>
        short Copies { get; set; }

        bool EndNote { get; set; }

        DisplayCellErrorType CellError { get; set; }
    }
}