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
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.Util
{
    /// <summary>
    /// <para>
    /// Program elements annotated [Removal] track the earliest final release
    /// when a deprecated feature will be removed. This is an internal decoration:
    /// a feature may be removed in a release earlier or later than the release
    /// number specified by this annotation.
    /// </para>
    /// <para>
    /// The POI project policy is to deprecate an element for 2 final releases
    /// before removing. This annotation exists to make it easier to follow up on the
    /// second step of the two-step deprecate and remove process.
    /// </para>
    /// <para>
    /// A deprecated feature may be removed in nightly and beta releases prior
    /// to the final release for which it is eligible, but may be removed later for
    /// various reasons. If it is known in advance that the feature will not be
    /// removed in the n+2 release, a later version should be specified by this
    /// annotation. The annotation version number should not include beta
    /// </para>
    /// <para>
    /// For example, a feature with a <c>deprecated POI 3.15 beta 3</c>
    /// is deprecated in POI 3.15 and 3.16 and becomes eligible for deletion during
    /// the POI 3.17 release series, and may be deleted immediately After POI 3.16 is
    /// released. This would be annotated <c>[Removal(version="3.17")]</c>.
    /// </para>
    /// </summary>
    /// <remarks>
    /// @since POI-3.15 beta 3
    /// </remarks>
    public class RemovalAttribute : Attribute
    {
        /// <summary>
        /// <para>
        /// The POI version when this feature may be removed.
        /// </para>
        /// <para>
        /// To ensure that the version number can be compared to the current version
        /// and a unit test can generate a warning if a removal-eligible feature has
        /// not been removed yet, the version number should adhere to the following format:
        /// </para>
        /// <para>
        /// Format: "(?&lt;major>\d+)\.(?&lt;minor>\d+)"
        /// </para>
        /// <para>
        /// Example: "3.15"
        /// </para>
        /// </summary>
        public string Version { get; set; } = string.Empty;
    }
}
