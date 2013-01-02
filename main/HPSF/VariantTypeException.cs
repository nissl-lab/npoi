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
    /// <summary>
    /// This exception is thrown if HPSF encounters a problem with a variant type.
    /// Concrete subclasses specifiy the problem further.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2004-06-21
    /// </summary>
    [Serializable]
    public abstract class VariantTypeException : HPSFException
    {
        private Object value;

        private long variantType;



        /// <summary>
        /// Initializes a new instance of the <see cref="VariantTypeException"/> class.
        /// </summary>
        /// <param name="variantType">The variant type causing the problem</param>
        /// <param name="value">The value who's variant type causes the problem</param>
        /// <param name="msg">A message text describing the problem</param>
        public VariantTypeException(long variantType, Object value,
                                    String msg):base(msg)
        {
            
            this.variantType = variantType;
            this.value = value;
        }

        /// <summary>
        /// Gets the offending variant type
        /// </summary>
        /// <value>the offending variant type.</value>
        public long VariantType
        {
            get { return variantType; }
        }

        /// <summary>
        /// Returns the value who's variant type caused the problem.
        /// </summary>
        /// <value>the value who's variant type caused the problem.</value>
        public Object Value
        {
            get{return value;}
        }

    }
}