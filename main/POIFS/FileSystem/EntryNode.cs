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

using NPOI.POIFS.Properties;

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// Abstract implementation of Entry
    /// Extending classes should override isDocument() or isDirectory(), as
    /// appropriate
    /// Extending classes must override isDeleteOK()
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    [Serializable]
    public abstract class EntryNode : Entry
    {

        // the DocumentProperty backing this object
        protected Property _property;

        // this object's parent Entry
        protected DirectoryNode _parent;


        protected EntryNode()
            : this(null, null)
        {
        }

        /// <summary>
        /// Create a DocumentNode. ThIs method Is not public by design; it
        /// Is intended strictly for the internal use of extending classes
        /// </summary>
        /// <param name="property">the Property for this Entry</param>
        /// <param name="parent">the parent of this entry</param>
        protected EntryNode(Property property, DirectoryNode parent)
        {
            _property = property;
            _parent = parent;
        }

        /// <summary>
        /// grant access to the property
        /// </summary>
        /// <value>the property backing this entry</value>
        public Property Property
        {
            get { return _property; }
        }

        /// <summary>
        /// Is this the root of the tree?
        /// </summary>
        /// <value><c>true</c> if this instance is root; otherwise, <c>false</c>.</value>
        protected bool IsRoot
        {
            get
            {
                // only the root Entry has no parent ...
                return (_parent == null);
            }
        }

        /// <summary>
        /// extensions use this method to verify internal rules regarding
        /// deletion of the underlying store.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if it's ok to Delete the underlying store; otherwise, <c>false</c>.
        /// </value>
        protected abstract bool IsDeleteOK { get; }

        /// <summary>
        /// Get the name of the Entry
        /// </summary>
        /// <value>The name.</value>
        /// Get the name of the Entry
        /// @return name
        public String Name
        {
            get
            {
                return _property.Name;
            }
        }

        /// <summary>
        /// Is this a DirectoryEntry?
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the Entry Is a DirectoryEntry; otherwise, <c>false</c>.
        /// </value>
        public virtual bool IsDirectoryEntry
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Is this a DocumentEntry?
        /// </summary>
        /// <value>
        /// 	<c>true</c> if the Entry Is a DocumentEntry; otherwise, <c>false</c>.
        /// </value>
        public virtual bool IsDocumentEntry
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Get this Entry's parent (the DocumentEntry that owns this
        /// Entry). All Entry objects, except the root Entry, has a parent.
        /// </summary>
        /// <value>this Entry's parent; null iff this Is the root Entry</value>
        public DirectoryEntry Parent
        {
            get
            {
                return _parent;
            }
        }

        /// <summary>
        /// Delete this Entry. ThIs operation should succeed, but there are
        /// special circumstances when it will not:
        /// If this Entry Is the root of the Entry tree, it cannot be
        /// deleted, as there Is no way to Create another one.
        /// If this Entry Is a directory, it cannot be deleted unless it Is
        /// empty.
        /// </summary>
        /// <returns>
        /// true if the Entry was successfully deleted, else false
        /// </returns>
        public bool Delete()
        {
            bool rval = false;

            if ((!IsRoot) && IsDeleteOK)
            {
                rval = _parent.DeleteEntry(this);
            }
            return rval;
        }

        /// <summary>
        /// Rename this Entry. ThIs operation will fail if:
        /// There Is a sibling Entry (i.e., an Entry whose parent Is the
        /// same as this Entry's parent) with the same name.
        /// ThIs Entry Is the root of the Entry tree. Its name Is dictated
        /// by the Filesystem and many not be Changed.
        /// </summary>
        /// <param name="newName">the new name for this Entry</param>
        /// <returns>
        /// true if the operation succeeded, else false
        /// </returns>
        public bool RenameTo(String newName)
        {
            bool rval = false;

            if (!IsRoot)
            {
                rval = _parent.ChangeName(Name, newName);
            }
            return rval;
        }

        public override string ToString()
        {
            return this.Name;
        }
    }
}
