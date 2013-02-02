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

using System.Collections.Generic;

namespace NPOI.POIFS.Properties
{
    /// <summary>
    /// Behavior for parent (directory) properties
    /// @author Marc Johnson27591@hotmail.com
    /// </summary>
    public interface Parent:Child
    {
        /// <summary>
        /// Get an iterator over the children of this Parent
        /// all elements are instances of Property.
        /// </summary>
        /// <returns></returns>
        IEnumerator<Property> Children { get; }
        /// <summary>
        /// Add a new child to the collection of children
        /// </summary>
        /// <param name="property">the new child to be added; must not be null</param>
        void AddChild(Property property);
        /// <summary>
        /// Sets the previous child.
        /// </summary>
        new Child PreviousChild { get; set; }
        /// <summary>
        /// Sets the next child.
        /// </summary>
        new Child NextChild { get; set; }
    }
}
