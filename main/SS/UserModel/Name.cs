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
    /**
     * Represents a defined name for a range of cells.
     * A name is a meaningful shorthand that makes it easier to understand the purpose of a
     * cell reference, constant or a formula.
     */
    public interface IName
    {

        /**
         * Get the sheets name which this named range is referenced to
         *
         * @return sheet name, which this named range refered to
         */
        String SheetName { get; }

        /**
         * Gets the name of the named range
         *
         * @return named range name
         */
        String NameName { get; set; }


        /**
         * Returns the formula that the name is defined to refer to.
         *
         * @return the reference for this name, <code>null</code> if it has not been set yet. Never empty string
         * @see #SetRefersToFormula(String)
         */
        String RefersToFormula { get; set; }

        /**
         * Checks if this name is a function name
         *
         * @return true if this name is a function name
         */
        bool IsFunctionName { get; }

        /**
         * Checks if this name points to a cell that no longer exists
         *
         * @return <c>true</c> if the name refers to a deleted cell, <c>false</c> otherwise
         */
        bool IsDeleted { get; }

        /**
         * Returns the sheet index this name applies to.
         *
         * @return the sheet index this name applies to, -1 if this name applies to the entire workbook
         */
        int SheetIndex { get; set; }

        /**
         * Returns the comment the user provided when the name was Created.
         *
         * @return the user comment for this named range
         */
        String Comment { get; set; }
        /**
         * Indicates that the defined name refers to a user-defined function.
         * This attribute is used when there is an add-in or other code project associated with the file.
         *
         * @param value <c>true</c> indicates the name refers to a function.
         */
        void SetFunction(bool value);
    }

}