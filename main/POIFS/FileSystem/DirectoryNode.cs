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
using System.Collections;


using NPOI.POIFS.Properties;
using NPOI.POIFS.Dev;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.EventFileSystem;
using NPOI.Util;
using NPOI.Util.Collections;

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// Simple implementation of DirectoryEntry
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    [Serializable]
    public class DirectoryNode : EntryNode, DirectoryEntry, POIFSViewable, IEnumerable<Entry>
    {
        // Map of Entry instances, keyed by their names
        private Dictionary<string, Entry> _byname;

        private List<Entry> _entries;

        // the POIFSFileSystem we belong to
        private POIFSFileSystem _oFilesSystem;

        private NPOIFSFileSystem _nFilesSystem;

        // the path described by this document
        private POIFSDocumentPath _path;


        public DirectoryNode(DirectoryProperty property,
                        POIFSFileSystem fileSystem,
                        DirectoryNode parent)
            : this(property, parent, fileSystem, (NPOIFSFileSystem)null)
        {
        }

        /// <summary>
        /// Create a DirectoryNode. This method Is not public by design; it
        /// Is intended strictly for the internal use of this package
        /// </summary>
        /// <param name="property">the DirectoryProperty for this DirectoryEntry</param>
        /// <param name="nFileSystem">the POIFSFileSystem we belong to</param>
        /// <param name="parent">the parent of this entry</param>
        public DirectoryNode(DirectoryProperty property,
                NPOIFSFileSystem nFileSystem,
                DirectoryNode parent)
            : this(property, parent, (POIFSFileSystem)null, nFileSystem)
        {
        }

        private DirectoryNode(DirectoryProperty property,
                        DirectoryNode parent,
                        POIFSFileSystem oFileSystem,
                        NPOIFSFileSystem nFileSystem)
            : base(property, parent)
        {
            this._oFilesSystem = oFileSystem;
            this._nFilesSystem = nFileSystem;

            if (parent == null)
                _path = new POIFSDocumentPath();
            else
            {
                _path = new POIFSDocumentPath(parent._path, new string[] { property.Name });
            }

            _byname = new Dictionary<string, Entry>();
            _entries = new List<Entry>();
            IEnumerator<Property> iter = property.Children;

            while (iter.MoveNext())
            {
                Property child = iter.Current;
                Entry childNode = null;

                if (child.IsDirectory)
                {
                    DirectoryProperty childDir = (DirectoryProperty)child;
                    if (_oFilesSystem != null)
                    {
                        childNode = new DirectoryNode(childDir, _oFilesSystem, this);
                    }
                    else
                    {
                        childNode = new DirectoryNode(childDir, _nFilesSystem, this);
                    }
                }
                else
                {
                    childNode = new DocumentNode((DocumentProperty)child, this);
                }
                _entries.Add(childNode);
                _byname.Add(childNode.Name, childNode);
            }
        }
        
        /// <summary>
        /// open a document in the directory's entry's list of entries
        /// </summary>
        /// <param name="documentName">the name of the document to be opened</param>
        /// <returns>a newly opened DocumentStream</returns>
        public DocumentInputStream CreatePOIFSDocumentReader(
                String documentName)
        {
            Entry document = GetEntry(documentName);

            if (!document.IsDocumentEntry)
            {
                throw new IOException("Entry '" + documentName
                                      + "' Is not a DocumentEntry");
            }
            return new DocumentInputStream((DocumentEntry)document);
        }

        /// <summary>
        /// Create a new DocumentEntry; the data will be provided later
        /// </summary>
        /// <param name="document">the name of the new documentEntry</param>
        /// <returns>the new DocumentEntry</returns>
        public DocumentEntry CreateDocument(POIFSDocument document)
        {
            DocumentProperty property = document.DocumentProperty;
            DocumentNode     rval     = new DocumentNode(property, this);

            ((DirectoryProperty)Property).AddChild(property);
            _oFilesSystem.AddDocument(document);

            _entries.Add(rval);
            _byname.Add(property.Name, rval);

            return rval;
        }

        /// <summary>
        /// Change a contained Entry's name
        /// </summary>
        /// <param name="oldName">the original name</param>
        /// <param name="newName">the new name</param>
        /// <returns>true if the operation succeeded, else false</returns>
        public bool ChangeName(String oldName, String newName)
        {
            bool   rval  = false;
            EntryNode child = (EntryNode)_byname[oldName];

            if (child != null)
            {
                rval = ((DirectoryProperty)Property)
                    .ChangeName(child.Property, newName);
                if (rval)
                {
                    _byname.Remove(oldName);
                    _byname[child.Property.Name] = child;
                }
            }
            return rval;
        }

        /// <summary>
        /// Deletes the entry.
        /// </summary>
        /// <param name="entry">the EntryNode to be Deleted</param>
        /// <returns>true if the entry was Deleted, else false</returns>
        public bool DeleteEntry(EntryNode entry)
        {
            bool rval =
                ((DirectoryProperty)Property)
                    .DeleteChild(entry.Property);

            if (rval)
            {
                _entries.Remove(entry);
                _byname.Remove(entry.Name);

                if (_oFilesSystem != null)
                {
                    _oFilesSystem.Remove(entry);
                }
                else
                {
                    try
                    {
                        _nFilesSystem.Remove(entry);
                    }
                    catch (IOException)
                    {
                        // TODO Work out how to report this, given we can't change the method signature...
                    }
                }
            }
            return rval;
        }

        /// <summary>
        /// Gets the path.
        /// </summary>
        /// <value>this directory's path representation</value>
        public POIFSDocumentPath Path
        {
            get { return _path; }
        }

        public POIFSFileSystem FileSystem
        {
            get { return _oFilesSystem; }
        }

        public NPOIFSFileSystem NFileSystem
        {
            get { return _nFilesSystem; }
        }

       /// <summary>
        /// get an iterator of the Entry instances contained directly in
        /// this instance (in other words, children only; no grandchildren
        /// etc.)
        /// </summary>
        /// <value>
        /// The entries.never null, but hasNext() may return false
        /// immediately (i.e., this DirectoryEntry is empty). All
        /// objects retrieved by next() are guaranteed to be
        /// implementations of Entry.
        /// </value>
        public IEnumerator<Entry> Entries
        {
            get { return _entries.GetEnumerator(); }
        }
        internal Entry GetEntry(int index)
        {
            return _entries[index];
        }

        /**
         * get the names of all the Entries contained directly in this
         * instance (in other words, names of children only; no grandchildren
         * etc).
         *
         * @return the names of all the entries that may be retrieved with
         *         getEntry(String), which may be empty (if this 
         *         DirectoryEntry is empty)
         */
        public List<String> EntryNames
        {
            get
            {
                return new List<string>(_byname.Keys);
            }
        }

        /// <summary>
        /// is this DirectoryEntry empty?
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance contains no Entry instances; otherwise, <c>false</c>.
        /// </value>
        public bool IsEmpty
        {
            get { return _entries.Count == 0; }
        }

        /// <summary>
        /// find out how many Entry instances are contained directly within
        /// this DirectoryEntry
        /// </summary>
        /// <value>
        /// number of immediately (no grandchildren etc.) contained
        /// Entry instances
        /// </value>
        public int EntryCount
        {
            get { return _entries.Count; }
        }


        public bool HasEntry(String name)
        {
            return name != null && _byname.ContainsKey(name);
        }

        /// <summary>
        /// get a specified Entry by name
        /// </summary>
        /// <param name="name">the name of the Entry to obtain.</param>
        /// <returns>
        /// the specified Entry, if it is directly contained in
        /// this DirectoryEntry
        /// </returns>
        public Entry GetEntry(String name)
        {
            Entry rval = null;

            if (name != null)
            {
                try
                {
                    rval = (Entry)_byname[name];
                }
                catch (KeyNotFoundException)
                {
                    throw new FileNotFoundException("no such entry: \"" + name + "\"");
                }
            }
            if (rval == null)
            {

                // either a null name was given, or there Is no such name
                throw new FileNotFoundException("no such entry: \"" + name + "\"");
            }
            return rval;
        }


        //public DocumentReader CreateDocumentReader(string documentName)
        //{
        //    try
        //    {
        //        return CreateDocumentReader(GetEntry(documentName));
        //    }
        //    catch(IOException ex)
        //    {
        //        throw ex;
        //    }
        //}

        //public DocumentReader CreateDocumentReader(Entry document)
        //{
        //    if (!document.IsDirectoryEntry)
        //    {
        //        throw new IOException("Entry '" + document.Name + "' is not a DocumentEntry");
        //    }

        //    DocumentEntry entry = (DocumentEntry)document;
        //    return new DocumentReader(entry);
        //}

        public DocumentInputStream CreateDocumentInputStream(Entry document)
        {
            if (!document.IsDocumentEntry)
            {
                throw new IOException("Entry '" + document.Name
                                    + "' is not a DocumentEntry");
            }

            DocumentEntry entry = (DocumentEntry)document;
            return new DocumentInputStream(entry);
        }

        public DocumentInputStream CreateDocumentInputStream(string documentName)
        {
            return CreateDocumentInputStream(GetEntry(documentName));
        }


        public DocumentEntry CreateDocument(NPOIFSDocument document)
        {
            try
            {
                DocumentProperty property = document.DocumentProperty;
                DocumentNode rval = new DocumentNode(property, this);

                ((DirectoryProperty)Property).AddChild(property);

                _nFilesSystem.AddDocument(document);

                _entries.Add(rval);
                _byname[property.Name] = rval;

                return rval;

            }
            catch (IOException ex)
        {
                throw ex;
            }
        }


        /// <summary>
        /// Create a new DirectoryEntry
        /// </summary>
        /// <param name="name">the name of the new DirectoryEntry</param>
        /// <returns>the name of the new DirectoryEntry</returns>
        public DirectoryEntry CreateDirectory(String name)
        {
            DirectoryProperty property = new DirectoryProperty(name);
            DirectoryNode rval;

            if (_oFilesSystem != null)
            {
                rval = new DirectoryNode(property, _oFilesSystem, this);
                _oFilesSystem.AddDirectory(property);
            }
            else
            {
                rval = new DirectoryNode(property, _nFilesSystem, this);
                _nFilesSystem.AddDirectory(property);
            }

            ((DirectoryProperty)Property).AddChild(property);
            _entries.Add(rval);
            _byname[name] = rval;

            return rval;
        }


        /// <summary>
        /// Gets or Sets the storage clsid for the directory entry
        /// </summary>
        /// <value>The storage ClassID.</value>
        public ClassID StorageClsid
        {
            set{
                this.Property.StorageClsid=value;
            }
            get
            {
                return this.Property.StorageClsid;
            }
        }

        /// <summary>
        /// Is this a DirectoryEntry?
        /// </summary>
        /// <value>true if the Entry Is a DirectoryEntry, else false</value>
        public override bool IsDirectoryEntry
        {
            get { return true; }
        }

        /// <summary>
        /// extensions use this method to verify internal rules regarding
        /// deletion of the underlying store.
        /// </summary>
        /// <value> true if it's ok to Delete the underlying store, else
        /// false</value>
        protected override bool IsDeleteOK
        {
            // if this directory Is empty, we can Delete it
            get
            {
                return IsEmpty;
            }
        }

        public DocumentEntry CreateDocument(string name, Stream stream)
        {
            try
            {
                if (_nFilesSystem != null)
                {
                    return CreateDocument(new NPOIFSDocument(name, _nFilesSystem, stream));
                }
                else
                {
                    return CreateDocument(new POIFSDocument(name, stream));
                }
            }
            catch (IOException ex)
            {
                throw ex;
            }
        }


        public DocumentEntry CreateDocument(string name, int size, POIFSWriterListener writer)
        {
            if (_nFilesSystem != null)
            {
                return CreateDocument(new NPOIFSDocument(name, size, _nFilesSystem, writer));
            }
            else
            {
                return CreateDocument(new POIFSDocument(name, size, _path, writer));
            }
        }


        #region ViewableItertor interface
        /// <summary>
        /// Get an array of objects, some of which may implement POIFSViewable
        /// </summary>
        /// <value>an array of Object; may not be null, but may be empty</value>
        public Array ViewableArray
        {
            get
            {
                return new Object[0];
            }
        }

        /// <summary>
        /// Get an Iterator of objects, some of which may implement
        /// POIFSViewable
        /// </summary>
        /// <value>an Iterator; may not be null, but may have an empty
        /// back end store</value>
        public IEnumerator ViewableIterator
        {
            get
                {
                ArrayList components = new ArrayList();

                components.Add(Property);
                components.AddRange(this._entries);
                //components.Sort();
                return components.GetEnumerator();
            }
        }

        /// <summary>
        /// Give viewers a hint as to whether to call GetViewableArray or
        /// GetViewableIterator
        /// </summary>
        /// <value><c>true</c> if a viewer should call GetViewableArray; otherwise, <c>false</c>if
        /// a viewer should call GetViewableIterator</value>
        public bool PreferArray
        {
            get { return false; }
        }

        /// <summary>
        /// Provides a short description of the object, to be used when a
        /// POIFSViewable object has not provided its contents.
        /// </summary>
        /// <value>The short description.</value>
        public String ShortDescription
        {
            get { return Name; }
        }
        
        #endregion



        #region IEnumerable<Entry> Members

        public IEnumerator<Entry> GetEnumerator()
        {
            return _entries.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _entries.GetEnumerator();
        }

        #endregion




        public  bool CanRead
        {
            get { throw new System.NotImplementedException(); }
        }

        public  bool CanSeek
        {
            get { throw new System.NotImplementedException(); }
        }

        public  bool CanWrite
        {
            get { throw new System.NotImplementedException(); }
        }

        public  void Flush()
        {
            throw new System.NotImplementedException();
        }

        public  long Length
        {
            get { throw new System.NotImplementedException(); }
        }

        public  long Position
        {
            get
            {
                throw new System.NotImplementedException();
    }
            set
            {
                throw new System.NotImplementedException();
            }
        }

        public  int Read(byte[] buffer, int offset, int count)
        {
            throw new System.NotImplementedException();
        }

        public  long Seek(long offset, SeekOrigin origin)
        {
            throw new System.NotImplementedException();
        }

        public  void SetLength(long value)
        {
            throw new System.NotImplementedException();
        }




    }


}
