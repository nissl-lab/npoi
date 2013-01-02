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

        /**    
         * Returns the paper size.    
         * @return paper size    
         */
        short PaperSize { get; set; }

        /**    
         * Returns the scale.    
         * @return scale    
         */
        short Scale{ get; set; }

        /**    
         * Returns the page start.    
         * @return page start    
         */
        short PageStart{ get; set; }

        /**    
         * Returns the number of pages wide to fit sheet in.    
         * @return number of pages wide to fit sheet in    
         */
        short FitWidth{ get; set; }

        /**    
         * Returns the number of pages high to fit the sheet in.    
         * @return number of pages high to fit the sheet in    
         */
        short FitHeight{ get; set; }

        /**    
         * Returns the left to right print order.    
         * @return left to right print order    
         */
        bool LeftToRight{ get; set; }

        /**    
         * Returns the landscape mode.    
         * @return landscape mode    
         */
        bool Landscape{ get; set; }

        /**    
         * Returns the valid Settings.    
         * @return valid Settings    
         */
        bool ValidSettings{ get; set; }

        /**    
         * Returns the black and white Setting.    
         * @return black and white Setting    
         */
        bool NoColor{ get; set; }

        /**    
         * Returns the draft mode.    
         * @return draft mode    
         */
        bool Draft{ get; set; }

        /**    
         * Returns the print notes.    
         * @return print notes    
         */
        bool Notes{ get; set; }

        /**    
         * Returns the no orientation.    
         * @return no orientation    
         */
        bool NoOrientation{ get; set; }

        /**    
         * Returns the use page numbers.    
         * @return use page numbers    
         */
        bool UsePage{ get; set; }

        /**    
         * Returns the horizontal resolution.    
         * @return horizontal resolution    
         */
        short HResolution{ get; set; }

        /**    
         * Returns the vertical resolution.    
         * @return vertical resolution    
         */
        short VResolution{ get; set; }

        /**    
         * Returns the header margin.    
         * @return header margin    
         */
        double HeaderMargin{ get; set; }

        /**    
         * Returns the footer margin.    
         * @return footer margin    
         */
        double FooterMargin{ get; set; }

        /**    
         * Returns the number of copies.    
         * @return number of copies    
         */
        short Copies{ get; set; }

        bool EndNote { get; set; }

        DisplayCellErrorType CellError { get; set; }
    }
}