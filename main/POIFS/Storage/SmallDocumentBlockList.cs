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

namespace NPOI.POIFS.Storage
{
    /// <summary>
    /// A list of SmallDocumentBlocks instances, and methods to manage the list
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class SmallDocumentBlockList:BlockListImpl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SmallDocumentBlockList"/> class.
        /// </summary>
        /// <param name="blocks">a list of SmallDocumentBlock instances</param>
        public SmallDocumentBlockList(List<SmallDocumentBlock> blocks)
        {
            SetBlocks((ListManagedBlock[])blocks.ToArray());
        }
    }
}
