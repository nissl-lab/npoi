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
    /// Interface for classes providing differential style definitions, such as conditional format rules
    /// and table/pivot table styles.
    /// </summary>
    /// @since 3.17 beta 1
    public interface IDifferentialStyleProvider
    {
        /// <summary>
        /// border formatting object  if defined,  <c>null</c> otherwise
        /// </summary>
        IBorderFormatting BorderFormatting { get; }
        /// <summary>
        /// font formatting object  if defined,  <c>null</c> otherwise
        /// </summary>
        IFontFormatting FontFormatting { get; }
        /// <summary>
        /// pattern formatting object if defined, <c>null</c> otherwise
        /// </summary>
        IPatternFormatting PatternFormatting { get; }
        /// <summary>
        /// number format defined for this rule, or null if the cell default should be used
        /// </summary>
        ExcelNumberFormat NumberFormat { get; }
        int StripeSize { get; }
    }
}
