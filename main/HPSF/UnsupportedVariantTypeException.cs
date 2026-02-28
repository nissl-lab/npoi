/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS Is" BASIS,
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

namespace NPOI.HPSF
{
    using System;
    using NPOI.Util;

    /// <summary>
    /// This exception is thrown if HPSF encounters a variant type that isn't
    /// supported yet. Although a variant type is unsupported the value can still be
    /// retrieved using the {@link VariantTypeException#GetValue} method.
    /// Obviously this class should disappear some day.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2003-08-05
    /// </summary>
    [Serializable]
    public abstract class UnsupportedVariantTypeException : VariantTypeException
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="UnsupportedVariantTypeException"/> class.
        /// </summary>
        /// <param name="variantType">The unsupported variant type</param>
        /// <param name="value">The value who's variant type is not yet supported</param>
        public UnsupportedVariantTypeException(long variantType,
                                               Object value)
            : base(variantType, value,
                "HPSF does not yet support the variant type " + variantType +
                " (" + Variant.GetVariantName(variantType) + ", " +
                HexDump.ToHex(variantType) + "). If you want support for " +
                "this variant type in one of the next POI releases please " +
                "submit a request for enhancement (RFE) To " +
                "<http://issues.apache.org/bugzilla/>! Thank you!")
        {

        }



    }
}