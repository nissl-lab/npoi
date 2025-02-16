/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System.Collections.Generic;

namespace NPOI.HPSF
{
    using System;
    using System.IO;
    using System.Collections;
    using NPOI.Util;
    using NPOI.POIFS.FileSystem;


    /// <summary>
    /// Adds writing support To the {@link PropertySet} class.
    /// Please be aware that this class' functionality will be merged into the
    /// {@link PropertySet} class at a later time, so the API will Change.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2003-02-19
    /// </summary>
    [Obsolete("deprecated POI 3.16 - use PropertySet as base class instead")]
    public class MutablePropertySet : PropertySet
    {

        public MutablePropertySet()
        {

        }



        /// <summary>
        /// Initializes a new instance of the <see cref="MutablePropertySet"/> class.
        /// All nested elements, i.e.<c>Section</c>s and <c>Property</c> instances, will be their
        /// mutable counterparts in the new <c>MutablePropertySet</c>.
        /// </summary>
        /// <param name="ps">The property Set To copy</param>
        public MutablePropertySet(PropertySet ps)
            : base(ps)
        {

        }

    }
}