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

namespace NPOI.POIFS.Storage
{

    /// <summary>
    /// Interface for lists of blocks that are mapped by block allocation
    /// tables
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public interface BlockList
    {
        /// <summary>
        /// remove the specified block from the list
        /// </summary>
        /// <param name="index">the index of the specified block; if the index is
        /// out of range, that's ok</param>
        void Zap(int index);

        /// <summary>
        /// Remove and return the specified block from the list
        /// </summary>
        /// <param name="index">the index of the specified block</param>
        /// <returns>the specified block</returns>
        ListManagedBlock Remove(int index);

        /// <summary>
        /// get the blocks making up a particular stream in the list. The
        /// blocks are removed from the list.
        /// </summary>
        /// <param name="startBlock">the index of the first block in the stream</param>
        /// <param name="headerPropertiesStartBlock"></param>
        /// <returns>the stream as an array of correctly ordered blocks</returns>
        ListManagedBlock[] FetchBlocks(int startBlock,int headerPropertiesStartBlock);

        /// <summary>
        /// set the associated BlockAllocationTable
        /// </summary>
        /// <value>the associated BlockAllocationTable</value>
        BlockAllocationTableReader BAT { set; }

        int BlockCount();
    }
}
