/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
    /// <para>
    /// Ordered list of table style elements, for both data tables and pivot tables.
    /// Some elements only apply to pivot tables, but any style definition can omit any number,
    /// so having them in one list should not be a problem.
    /// </para>
    /// <para>
    /// The order is the specification order of application, with later elements overriding previous
    /// ones, if style properties conflict.
    /// </para>
    /// <para>
    /// Processing could iterate bottom-up if looking for specific properties, and stop when the
    /// first style is found defining a value for that property.
    /// </para>
    /// <para>
    /// Enum names match the OOXML spec values exactly, so <see cref="valueOf(String)" /> will work.
    /// </para>
    /// </summary>
    /// @since 3.17 beta 1
    public enum TableStyleType
    {
        wholeTable,
        pageFieldLabels,// pivot only
        pageFieldValues,// pivot only
        firstColumnStripe,
        secondColumnStripe,
        firstRowStripe,
        secondRowStripe,
        lastColumn,
        firstColumn,
        headerRow,
        totalRow,
        firstHeaderCell,
        lastHeaderCell,
        firstTotalCell,
        lastTotalCell,
        /* these are for pivot tables only */
        /***/
        firstSubtotalColumn,
        /***/
        secondSubtotalColumn,
        /***/
        thirdSubtotalColumn,
        /***/
        blankRow,
        /***/
        firstSubtotalRow,
        /***/
        secondSubtotalRow,
        /***/
        thirdSubtotalRow,
        /***/
        firstColumnSubheading,
        /***/
        secondColumnSubheading,
        /***/
        thirdColumnSubheading,
        /***/
        firstRowSubheading,
        /***/
        secondRowSubheading,
        /***/
        thirdRowSubheading
    }
}
