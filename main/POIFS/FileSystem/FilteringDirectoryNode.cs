/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using NPOI.Util.Collections;
using System;
using System.Collections.Generic;
using System.IO;

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// A DirectoryEntry filter, which exposes another  DirectoryEntry less certain parts.
    /// This is typically used when copying or comparing  Filesystems.
    /// </summary>
    public class FilteringDirectoryNode : DirectoryEntry
    {

        private List<String> excludes;
        private Dictionary<String, List<String>> childExcludes;
        private DirectoryEntry directory;
        /// <summary>
        /// Creates a filter round the specified directory, which will exclude entries such as 
        /// "MyNode" and "MyDir/IgnoreNode". The excludes can stretch into children, if they contain a /.
        /// </summary>
        /// <param name="directory">The Directory to filter</param>
        /// <param name="excludes">The Entries to exclude</param>
        public FilteringDirectoryNode(DirectoryEntry directory, ICollection<String> excludes)
        {
            this.directory = directory;

            // Process the excludes
            this.excludes = new List<String>();
            this.childExcludes = new Dictionary<String, List<String>>();
            foreach (String excl in excludes)
            {
                int splitAt = excl.IndexOf('/');
                if (splitAt == -1)
                {
                    // Applies to us
                    this.excludes.Add(excl);
                }
                else
                {
                    // Applies to a child
                    String child = excl.Substring(0, splitAt);
                    String childExcl = excl.Substring(splitAt + 1);
                    if (!this.childExcludes.ContainsKey(child))
                    {
                        this.childExcludes.Add(child, new List<String>());
                    }
                    this.childExcludes[child].Add(childExcl);
                }
            }
        }
        #region DirectoryEntry 成员

        public IEnumerator<Entry> Entries
        {
            get { return GetEntries(); }
        }

        public List<String> EntryNames
        {
            get
            {
                List<String> names = new List<String>();
                foreach (String name in directory.EntryNames)
                {
                    if (!excludes.Contains(name))
                    {
                        names.Add(name);
                    }
                }
                return names;
            }
        }

        public bool IsEmpty
        {
            get { return EntryCount == 0; }
        }
        public bool HasEntry(String name)
        {
            if (excludes.Contains(name))
            {
                return false;
            }
            return directory.HasEntry(name);
        }
        public int EntryCount
        {
            get
            {
                int size = directory.EntryCount;
                foreach (String excl in excludes)
                {
                    if (directory.HasEntry(excl))
                    {
                        size--;
                    }
                }
                return size;
            }
        }

        public IEnumerator<Entry> GetEntries()
        {
            return new FilteringIterator(this); ;
        }
        public Entry GetEntry(String name)
        {
            if (excludes.Contains(name))
            {
                throw new FileNotFoundException(name);
            }

            Entry entry = directory.GetEntry(name);
            return WrapEntry(entry);
        }
        private Entry WrapEntry(Entry entry)
        {
            String name = entry.Name;
            if (childExcludes.ContainsKey(name) && entry is DirectoryEntry)
            {
                return new FilteringDirectoryNode(
                      (DirectoryEntry)entry, childExcludes[name]);
            }
            return entry;
        }
        public DocumentEntry CreateDocument(string name, System.IO.Stream stream)
        {
            return directory.CreateDocument(name, stream);
        }

        public DocumentEntry CreateDocument(string name, int size, EventFileSystem.POIFSWriterListener writer)
        {
            return directory.CreateDocument(name, size, writer);
        }

        public DirectoryEntry CreateDirectory(string name)
        {
            return directory.CreateDirectory(name);
        }

        public Util.ClassID StorageClsid
        {
            get
            {
                return directory.StorageClsid;
            }
            set
            {
                directory.StorageClsid = value;
            }
        }

        #endregion

        #region Entry 成员

        public string Name
        {
            get { return directory.Name; }
        }

        public bool IsDirectoryEntry
        {
            get { return true; }
        }

        public bool IsDocumentEntry
        {
            get { return false; }
        }

        public DirectoryEntry Parent
        {
            get { return directory.Parent; }
        }

        public bool Delete()
        {
            return directory.Delete();
        }

        public bool RenameTo(string newName)
        {
            return directory.RenameTo(newName);
        }

        #endregion

        #region IEnumerable<Entry> 成员

        public IEnumerator<Entry> GetEnumerator()
        {
            return new FilteringIterator(this);
        }

        #endregion

        #region IEnumerable 成员

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new FilteringIterator(this); 
        }

        #endregion
        private class FilteringIterator : IEnumerator<Entry>
        {
            private IEnumerator<Entry> parent;
            private Entry next;
            private DirectoryEntry directory;
            private FilteringDirectoryNode filtering;
            public FilteringIterator(FilteringDirectoryNode filtering)
            {
                this.filtering = filtering;
                this.directory = filtering.directory;
                parent = directory.Entries;
                //LocateNext();
            }
            //private void LocateNext()
            //{
            //    next = null;
            //    Entry e;
            //    while (parent.MoveNext() && next == null)
            //    {
            //        e = parent.Current;
            //        if (!filtering.excludes.Contains(e.Name))
            //        {
            //            next = filtering.WrapEntry(e);
            //        }
            //    }
            //}

            //public bool HasNext()
            //{
            //    return (next != null);
            //}

            //public Entry Next()
            //{
            //    Entry e = next;
            //    LocateNext();
            //    return e;
            //}

            public void Remove()
            {
                throw new InvalidOperationException("Remove not supported");
            }

            #region IEnumerator<Entry> 成员

            public Entry Current
            {
                get { return next; }
            }

            #endregion

            #region IDisposable 成员

            public void Dispose()
            {
            }

            #endregion

            #region IEnumerator 成员

            object System.Collections.IEnumerator.Current
            {
                get { return next; }
            }

            public bool MoveNext()
            {
                next = null;
                Entry e;
                while (parent.MoveNext())
                {
                    e = parent.Current;
                    if (!filtering.excludes.Contains(e.Name))
                    {
                        next = filtering.WrapEntry(e);
                        break;
                    }
                }
                return (next != null);
                //throw new NotImplementedException();
            }

            public void Reset()
            {
                throw new NotImplementedException();
            }

            #endregion
        }

        
    }
}
