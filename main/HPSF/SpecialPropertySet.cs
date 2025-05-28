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

namespace NPOI.HPSF

{

    /// <summary>
    /// <para>
    /// Interface for the convenience classes {@link SummaryInformation}
    /// </para>
    /// <para>
    /// and {@link DocumentSummaryInformation}.
    /// </para>
    /// <para>
    /// This used to be an abstract class to support late loading
    /// of the SummaryInformation classes, as their concrete instance can
    /// only be determined After the PropertySet has been loaded.
    /// </para>
    /// </summary>

    // @Removal(version="3.18")
    [Obsolete("deprecated POI 3.16 - use PropertySet as base class instead")]
    public class SpecialPropertySet : MutablePropertySet
    {
        public SpecialPropertySet()
        {
        }

        public SpecialPropertySet(PropertySet ps) : base(ps)
        {

            ;
        }
    }
}

