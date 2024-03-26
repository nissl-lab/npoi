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
    /// Represents a defined name for a range of cells.
    /// A name is a meaningful shorthand that makes it easier to understand the purpose of a
    /// cell reference, constant or a formula.
    /// </summary>
    public interface IName
    {

        /// <summary>
        /// Get the sheets name which this named range is referenced to
        /// </summary>
        /// <returns>sheet name, which this named range refered to</returns>
        String SheetName { get; }

        /// <summary>
        /// Gets the name of the named range
        /// </summary>
        /// <returns>named range name</returns>
        String NameName { get; set; }


        /// <summary>
        /// Returns the formula that the name is defined to refer to.
        /// </summary>
        /// <returns>the reference for this name, <c>null</c> if it has not been set yet. Never empty string</returns>
        /// <see cref="SetRefersToFormula(String)" />
        String RefersToFormula { get; set; }

        /// <summary>
        /// Checks if this name is a function name
        /// </summary>
        /// <returns>true if this name is a function name</returns>
        bool IsFunctionName { get; }

        /// <summary>
        /// Checks if this name points to a cell that no longer exists
        /// </summary>
        /// <returns><c>true</c> if the name refers to a deleted cell, <c>false</c> otherwise</returns>
        bool IsDeleted { get; }

        /// <summary>
        /// Returns the sheet index this name applies to.
        /// </summary>
        /// <returns>the sheet index this name applies to, -1 if this name applies to the entire workbook</returns>
        int SheetIndex { get; set; }

        /// <summary>
        /// Returns the comment the user provided when the name was Created.
        /// </summary>
        /// <returns>the user comment for this named range</returns>
        String Comment { get; set; }
        /// <summary>
        /// Indicates that the defined name refers to a user-defined function.
        /// This attribute is used when there is an add-in or other code project associated with the file.
        /// </summary>
        /// <param name="value"><c>true</c> indicates the name refers to a function.</param>
        void SetFunction(bool value);
    }

}