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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

namespace NPOI.POIFS.Properties
{
    /// <summary>
    /// This interface defines methods for finding and setting sibling
    /// Property instances
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public interface Child
    {
        /// <summary>
        /// Gets or sets the previous child.
        /// </summary>
        /// <value>The previous child.</value>
        Child PreviousChild { get; set; }


        /// <summary>
        /// Gets or sets the next child.
        /// </summary>
        /// <value>The next child.</value>
        Child NextChild { get; set; }
    }
}
