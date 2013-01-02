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
using System;
using System.Collections;

namespace NPOI.Util.Collections
{
    /// <summary>
    /// This interface comes from Java
    /// </summary>
	public interface ISet
		: ICollection
	{
        /// <summary>
        /// Adds the specified o.
        /// </summary>
        /// <param name="o">The o.</param>
		void Add(object o);
        /// <summary>
        /// Determines whether [contains] [the specified o].
        /// </summary>
        /// <param name="o">The o.</param>
        /// <returns>
        /// 	<c>true</c> if [contains] [the specified o]; otherwise, <c>false</c>.
        /// </returns>
		bool Contains(object o);
        /// <summary>
        /// Removes the specified o.
        /// </summary>
        /// <param name="o">The o.</param>
		void Remove(object o);
        /// <summary>
        /// Removes all of the elements from this set (optional operation).
        /// The set will be empty after this call returns.
        /// </summary>
        void Clear();
	}
}
