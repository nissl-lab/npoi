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
    /// Common interface for NPOI.SS.UserModel.Header and NPOI.SS.UserModel.Footer
    /// </summary>
    public interface IHeaderFooter
    {
        /// <summary>
        /// Gets or sets the left side of the header or footer.
        /// </summary>
        /// <value>The string representing the left side.</value>
        String Left { get; set; }

        /// <summary>
        /// Gets or sets the center of the header or footer.
        /// </summary>
        /// <value>The string representing the center.</value>
        String Center { get; set; }

        /// <summary>
        /// Gets or sets the right side of the header or footer.
        /// </summary>
        /// <value>The string representing the right side.</value>
        String Right { get; set; }
    }
}

