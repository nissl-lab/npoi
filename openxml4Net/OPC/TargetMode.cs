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

namespace NPOI.OpenXml4Net.OPC
{
    /// <summary>
    /// Specifies whether the target of a PackageRelationship is inside or outside a
    /// Package.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 1.0
    /// </remarks>

    public enum TargetMode
    {
        /// <summary>
        /// The relationship references a resource that is external to the package. */
        /// </summary>
        Internal,
        /// <summary>
        /// The relationship references a part that is inside the package. */
        /// </summary>
        External
    }
}
