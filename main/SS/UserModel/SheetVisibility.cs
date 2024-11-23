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
    /// Specifies sheet visibility
    /// </summary>
    public enum SheetVisibility : int
    {
        /// <summary>
        /// Indicates the sheet is visible.
        /// </summary>
        Visible = 0,

        /// <summary>
        /// Indicates the book window is hidden, but can be shown by the user via the user interface.
        /// </summary>
        Hidden = 1,

        /// <summary>
        /// Indicates the sheet is hidden and cannot be shown in the user interface (UI).
        /// </summary>
        /// <remarks>
        /// In Excel this state is only available programmatically in VBA:
        /// ThisWorkbook.Sheets("MySheetName").Visible = xlSheetVeryHidden
        /// 
        /// </remarks>
        VeryHidden = 2
    }
}
