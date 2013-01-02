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
    /// This exception is thrown if one of the {@link PropertySet}'s
    /// convenience methods that require a single {@link Section} is called
    /// and the {@link PropertySet} does not contain exactly one {@link
    /// Section}.
    /// The constructors of this class are analogous To those of its
    /// superclass and documented there.
    /// @author Rainer Klute 
    /// <a href="mailto:klute@rainer-klute.de">&lt;klute@rainer-klute.de&gt;</a>
    /// @since 2002-02-09
    /// </summary>
    [Serializable]
    public class NoSingleSectionException : HPSFRuntimeException
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="NoSingleSectionException"/> class.
        /// </summary>
        public NoSingleSectionException():base()
        {
            
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="NoSingleSectionException"/> class.
        /// </summary>
        /// <param name="msg">The exception's message string</param>
        public NoSingleSectionException(String msg):base(msg)
        {
            
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="NoSingleSectionException"/> class.
        /// </summary>
        /// <param name="reason">This exception's underlying reason</param>
        public NoSingleSectionException(Exception reason):base(reason)
        {
            
        }


        /// <summary>
        /// Initializes a new instance of the <see cref="NoSingleSectionException"/> class.
        /// </summary>
        /// <param name="msg">The exception's message string</param>
        /// <param name="reason">This exception's underlying reason</param>
        public NoSingleSectionException(String msg, Exception reason):base(msg, reason)
        {
            
        }

    }
}