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

namespace NPOI.POIFS.Dev
{
    /// <summary>
    /// Interface for a drill-down viewable object. Such an object has
    /// content that may or may not be displayed, at the discretion of the
    /// viewer. The content is returned to the viewer as an array or as an
    /// Iterator, and the object provides a clue as to which technique the
    /// viewer should use to get its content.
    /// A POIFSViewable object is also expected to provide a short
    /// description of itself, that can be used by a viewer when the
    /// viewable object is collapsed.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public interface POIFSViewable
    {
        /// <summary>
        /// Provides a short description of the object to be used when a 
        /// POIFSViewable object has not provided its contents.
        /// </summary>
        /// <value><c>true</c> if [prefer array]; otherwise, <c>false</c>.</value>
        bool PreferArray { get; }
        /// <summary>
        /// Gets the short description.
        /// </summary>
        /// <value>The short description.</value>
        string ShortDescription { get; }
        /// <summary>
        /// Get an array of objects, some of which may implement POIFSViewable
        /// </summary>
        /// <value>The viewable array.</value>
        Array ViewableArray { get; }
        /// <summary>
        /// Give viewers a hint as to whether to call ViewableArray or ViewableIterator
        /// </summary>
        /// <value>The viewable iterator.</value>
        IEnumerator ViewableIterator { get; }

    }
}
