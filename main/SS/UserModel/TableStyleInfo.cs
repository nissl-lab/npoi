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
    /// style information for a specific table instance, referencing the document style
    /// and indicating which optional portions of the style to apply.
    /// </summary>
    /// @since 3.17 beta 1
    public interface ITableStyleInfo
    {
        /// <summary>
        /// true if alternating column styles should be applied
        /// </summary>
        bool IsShowColumnStripes { get; set; }
        /// <summary>
        /// return true if alternating row styles should be applied
        /// </summary>
        bool IsShowRowStripes { get; set; }
        /// <summary>
        /// return true if the distinct first column style should be applied
        /// </summary>
        bool IsShowFirstColumn { get; set; }

        /// <summary>
        /// return true if the distinct last column style should be applied
        /// </summary>
        bool IsShowLastColumn { get; set; }
        /// <summary>
        /// return the name of the style (may reference a built-in style)
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// style definition
        /// </summary>
        ITableStyle Style { get; }
    }
}
