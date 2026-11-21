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

using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    /// <summary>
    /// Data table style definition.  Includes style elements for various table components.
    /// Any number of style elements may be represented, and any cell may be styled by
    /// multiple elements.  The order of elements in <see cref="TableStyleType"/> defines precedence.
    /// </summary>
    /// @since 3.17 beta 1
    public interface ITableStyle
    {
        /// <summary>
        /// name (may be a built-in name)
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Some clients may care where in the table style list this definition came from, so we'll track it.
        /// The spec only references these by name, unlike Dxf records, which these definitions reference by index
        /// (XML definition order).  Nice of MS to be consistent when defining the ECMA standard.
        /// Use NPOI.xssf.UserModel.XSSFBuiltinTableStyle.IsBuiltinStyle(TableStyle) to determine whether the index is for a built-in style or explicit user style
        /// </summary>
        /// <return>index from NPOI.xssf.Model.StylesTable.GetExplicitTableStyle(String) or NPOI.xssf.UserModel.XSSFBuiltinTableStyle.ordinal()</return>
        int Index { get; }
        /// <summary>
        /// true if this is a built-in style defined in the OOXML specification, false if it is a user style
        /// </summary>
        bool IsBuiltin { get; }
        /// <summary>
        /// style definition for the given type, or null if not defined in this style.
        /// </summary>
        /// <param name="type"></param>
        /// <returns>style definition for the given type, or null if not defined in this style.</returns>
        IDifferentialStyleProvider GetStyle(TableStyleType type);
    }
}
