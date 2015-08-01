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
using System.Collections.Generic;
using System.IO;

using NPOI.Util;
using NPOI.POIFS.EventFileSystem;
using NPOI.Util.Collections;


namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// This interface defines methods specific to Directory objects
    /// managed by a Filesystem instance.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public interface DirectoryEntry : Entry, IEnumerable<Entry>
    {

        /// <summary>
        /// get an iterator of the Entry instances contained directly in
        /// this instance (in other words, children only; no grandchildren
        /// etc.)
        /// </summary>
        /// <value>The entries.never null, but hasNext() may return false
        /// immediately (i.e., this DirectoryEntry is empty). All
        /// objects retrieved by next() are guaranteed to be
        /// implementations of Entry.</value>
        IEnumerator<Entry> Entries { get; }

        /// <summary>
        /// get the names of all the Entries contained directly in this
        /// instance (in other words, names of children only; no grandchildren etc).
        /// </summary>
        /// <value>the names of all the entries that may be retrieved with
        /// getEntry(String), which may be empty (if this DirectoryEntry is empty
        /// </value>
        List<String> EntryNames { get; }

        /// <summary>
        ///is this DirectoryEntry empty?
        /// </summary>
        /// <value><c>true</c> if this instance contains no Entry instances; otherwise, <c>false</c>.</value>
        bool IsEmpty { get; }

        /// <summary>
        /// find out how many Entry instances are contained directly within
        /// this DirectoryEntry
        /// </summary>
        /// <value>number of immediately (no grandchildren etc.) contained
        /// Entry instances</value>
        int EntryCount{get;}

        /// <summary>
        /// get a specified Entry by name
        /// </summary>
        /// <param name="name">the name of the Entry to obtain.</param>
        /// <returns>the specified Entry, if it is directly contained in
        /// this DirectoryEntry</returns>
        Entry GetEntry(String name);

        /// <summary>
        /// Create a new DocumentEntry
        /// </summary>
        /// <param name="name">the name of the new DocumentEntry</param>
        /// <param name="stream">the Stream from which to Create the new DocumentEntry</param>
        /// <returns>the new DocumentEntry</returns>
        DocumentEntry CreateDocument(String name,
                                            Stream stream);
        // <summary>
        // Create a new DocumentEntry; the data will be provided later
        // </summary>
        // <param name="name">the name of the new DocumentEntry</param>
        // <param name="size">the size of the new DocumentEntry</param>
        // <returns>the new DocumentEntry</returns>
        //DocumentEntry CreateDocument(String name, int size);

        /// <summary>
        /// Create a new DocumentEntry; the data will be provided later
        /// </summary>
        /// <param name="name">the name of the new DocumentEntry</param>
        /// <param name="size">the size of the new DocumentEntry</param>
        /// <param name="writer">BeforeWriting event handler</param>
        /// <returns>the new DocumentEntry</returns>
        DocumentEntry CreateDocument(string name, int size, POIFSWriterListener writer);

        /// <summary>
        /// Create a new DirectoryEntry
        /// </summary>
        /// <param name="name">the name of the new DirectoryEntry</param>
        /// <returns>the name of the new DirectoryEntry</returns>
        DirectoryEntry CreateDirectory(String name);

        /// <summary>
        /// Gets or sets the storage ClassID.
        /// </summary>
        /// <value>The storage ClassID.</value>
        ClassID StorageClsid { get; set; }

        /// <summary>
        /// Checks if entry with specified name present
        /// </summary>
        /// <param name="name">entry name</param>
        /// <returns>true if have</returns>
        bool HasEntry(String name );
    }
}
