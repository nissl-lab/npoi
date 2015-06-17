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
    using System.IO;
    using NPOI.Util;


    /// <summary>
    /// Adds writing capability To the {@link Property} class.
    /// Please be aware that this class' functionality will be merged into the
    /// {@link Property} class at a later time, so the API will Change.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2003-08-03
    /// </summary>
    public class MutableProperty : Property
    {

        /// <summary>
        /// Creates an empty property. It must be Filled using the Set method To
        /// be usable.
        /// </summary>
        public MutableProperty()
        { }



        /// <summary>
        /// Initializes a new instance of the <see cref="MutableProperty"/> class.
        /// </summary>
        /// <param name="p">The property To copy.</param>
        public MutableProperty(Property p)
        {
            this.ID=p.ID;
            this.Type=p.Type;
            this.Value=p.Value;
        }


        /// <summary>
        /// Writes the property To an output stream.
        /// </summary>
        /// <param name="out1">The output stream To Write To.</param>
        /// <param name="codepage">The codepage To use for writing non-wide strings</param>
        /// <returns>the number of bytes written To the stream</returns>
        public int Write(Stream out1, int codepage)
        {
            int length = 0;
            long variantType = this.Type;

            /* Ensure that wide strings are written if the codepage is Unicode. */
            if (codepage == CodePageUtil.CP_UNICODE && variantType == Variant.VT_LPSTR)
                variantType = Variant.VT_LPWSTR;

            length += TypeWriter.WriteUIntToStream(out1, (uint)variantType);
            length += VariantSupport.Write(out1, variantType, this.Value, codepage);
            return length;
        }

    }
}