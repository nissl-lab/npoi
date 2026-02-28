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
    /// This exception is thrown when there is an illegal value Set in a
    /// {@link PropertySet}. For example, a {@link Variant#VT_BOOL} must
    /// have a value of <c>-1 (true)</c> or <c>0 (false)</c>.
    /// Any other value would trigger this exception. It supports a nested
    /// "reason" throwable, i.e. an exception that caused this one To be
    /// thrown.
    /// @author Drew Varner(Drew.Varner atDomain sc.edu)
    /// @since 2002-05-26
    /// </summary>
    [Serializable]
    public class IllegalPropertySetDataException : HPSFRuntimeException
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="IllegalPropertySetDataException"/> class.
        /// </summary>
        public IllegalPropertySetDataException()
            : base()
        {

        }



        /// <summary>
        /// Initializes a new instance of the <see cref="IllegalPropertySetDataException"/> class.
        /// </summary>
        /// <param name="msg">The exception's message string</param>
        public IllegalPropertySetDataException(String msg)
            : base(msg)
        {

        }



        /// <summary>
        /// Initializes a new instance of the <see cref="IllegalPropertySetDataException"/> class.
        /// </summary>
        /// <param name="reason">This exception's underlying reason</param>
        public IllegalPropertySetDataException(Exception reason)
            : base(reason)
        {

        }



        /// <summary>
        /// Initializes a new instance of the <see cref="IllegalPropertySetDataException"/> class.
        /// </summary>
        /// <param name="msg"> The exception's message string</param>
        /// <param name="reason">This exception's underlying reason</param>
        public IllegalPropertySetDataException(String msg,
                                               Exception reason)
            : base(msg, reason)
        {

        }

    }
}
