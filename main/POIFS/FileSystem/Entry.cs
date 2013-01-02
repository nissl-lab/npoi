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

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// This interface provides access to an object managed by a Filesystem
    /// instance. Entry objects are further divided into DocumentEntry and
    /// DirectoryEntry instances.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public interface Entry
    {
        /// <summary>
        /// Get the name of the Entry
        /// </summary>
        /// <value>The name.</value>
        String Name{get;}

        /// <summary>
        /// Is this a DirectoryEntry?
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the Entry Is a DirectoryEntry; otherwise, <c>false</c>.
        /// </value>
        bool IsDirectoryEntry{get;}

        /// <summary>
        /// Is this a DocumentEntry?
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the Entry Is a DocumentEntry; otherwise, <c>false</c>.
        /// </value>
        bool IsDocumentEntry { get; }

        /// <summary>
        /// Get this Entry's parent (the DirectoryEntry that owns this
        /// Entry). All Entry objects, except the root Entry, has a parent.
        /// </summary>
        /// <value>this Entry's parent; null iff this Is the root Entry</value>
        /// This property is moved to EntryNode
        DirectoryEntry Parent { get; }

        /// <summary>
        /// Delete this Entry. ThIs operation should succeed, but there are
        /// special circumstances when it will not:
        /// If this Entry Is the root of the Entry tree, it cannot be
        /// deleted, as there Is no way to Create another one.
        /// If this Entry Is a directory, it cannot be deleted unless it Is
        /// empty.
        /// </summary>
        /// <returns>true if the Entry was successfully deleted, else false</returns>
        bool Delete();

        /// <summary>
        /// Rename this Entry. ThIs operation will fail if:
        /// There Is a sibling Entry (i.e., an Entry whose parent Is the
        /// same as this Entry's parent) with the same name.
        /// ThIs Entry Is the root of the Entry tree. Its name Is dictated
        /// by the Filesystem and many not be Changed.
        /// </summary>
        /// <param name="newName">the new name for this Entry</param>
        /// <returns>true if the operation succeeded, else false</returns>
        bool RenameTo(String newName);
    }

}
