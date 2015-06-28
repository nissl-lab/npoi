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
using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;

namespace NPOI.POIFS.FileSystem
{
    public class EntryUtils
    {
        /**
     * Copies an Entry into a target POIFS directory, recursively
     */
        public static void CopyNodeRecursively(Entry entry, DirectoryEntry target)
        {
            // System.err.println("copyNodeRecursively called with "+entry.getName()+
            // ","+target.getName());
            DirectoryEntry newTarget = null;
            if (entry.IsDirectoryEntry)
            {
                DirectoryEntry dirEntry = (DirectoryEntry)entry;
                newTarget = target.CreateDirectory(entry.Name);
                newTarget.StorageClsid=(dirEntry.StorageClsid);
                IEnumerator<Entry> entries = dirEntry.Entries;

                while (entries.MoveNext())
                {
                    CopyNodeRecursively((Entry)entries.Current, newTarget);
                }
            }
            else
            {
                DocumentEntry dentry = (DocumentEntry)entry;
                DocumentInputStream dstream = new DocumentInputStream(dentry);
                target.CreateDocument(dentry.Name, dstream);
                dstream.Close();
            }
        }

        /**
         * Copies all the nodes from one POIFS Directory to another
         * 
         * @param sourceRoot
         *            is the source Directory to copy from
         * @param targetRoot
         *            is the target Directory to copy to
         */
        public static void CopyNodes(DirectoryEntry sourceRoot,
                DirectoryEntry targetRoot)
        {
            foreach (Entry entry in sourceRoot)
            {
                CopyNodeRecursively(entry, targetRoot);
            }
        }

        /**
         * Copies nodes from one Directory to the other minus the excepts
         * 
         * @param filteredSource The filtering source Directory to copy from
         * @param filteredTarget The filtering target Directory to copy to
         */
        public static void CopyNodes(FilteringDirectoryNode filteredSource,
                FilteringDirectoryNode filteredTarget)
        {
            // Nothing special here, just overloaded types to make the
            //  recommended new way to handle this clearer
            CopyNodes((DirectoryEntry)filteredSource, (DirectoryEntry)filteredTarget);
        }

        /**
         * Copies nodes from one Directory to the other minus the excepts
         * 
         * @param sourceRoot
         *            is the source Directory to copy from
         * @param targetRoot
         *            is the target Directory to copy to
         * @param excepts
         *            is a list of Strings specifying what nodes NOT to copy
         * @deprecated use {@link FilteringDirectoryNode} instead
         */
        [Obsolete]
        public static void CopyNodes(DirectoryEntry sourceRoot,
                DirectoryEntry targetRoot, List<String> excepts)
        {
            IEnumerator entries = sourceRoot.Entries;
            while (entries.MoveNext())
            {
                Entry entry = (Entry)entries.Current;
                if (!excepts.Contains(entry.Name))
                {
                    CopyNodeRecursively(entry, targetRoot);
                }
            }
        }

        /**
         * Copies all nodes from one POIFS to the other
         * 
         * @param source
         *            is the source POIFS to copy from
         * @param target
         *            is the target POIFS to copy to
         */
        public static void CopyNodes(POIFSFileSystem source,
                POIFSFileSystem target)
        {
            CopyNodes(source.Root, target.Root);
        }

        /**
         * Copies nodes from one POIFS to the other, minus the excepts.
         * This delegates the filtering work to {@link FilteringDirectoryNode},
         *  so excepts can be of the form "NodeToExclude" or
         *  "FilteringDirectory/ExcludedChildNode"
         * 
         * @param source is the source POIFS to copy from
         * @param target is the target POIFS to copy to
         * @param excepts is a list of Entry Names to be excluded from the copy
         */
        public static void CopyNodes(POIFSFileSystem source,
                POIFSFileSystem target, List<String> excepts)
        {
            CopyNodes(
                  new FilteringDirectoryNode(source.Root, excepts),
                  new FilteringDirectoryNode(target.Root, excepts)
            );
        }

        /**
         * Checks to see if the two Directories hold the same contents.
         * For this to be true, they must have entries with the same names,
         *  no entries in one but not the other, and the size+contents
         *  of each entry must match, and they must share names.
         * To exclude certain parts of the Directory from being checked,
         *  use a {@link FilteringDirectoryNode}
         */
        public static bool AreDirectoriesIdentical(DirectoryEntry dirA, DirectoryEntry dirB)
        {
            // First, check names
            if (!dirA.Name.Equals(dirB.Name))
            {
                return false;
            }

            // Next up, check they have the same number of children
            if (dirA.EntryCount != dirB.EntryCount)
            {
                return false;
            }

            // Next, check entries and their types/sizes
            Dictionary<String, int> aSizes = new Dictionary<String, int>();
            int isDirectory = -12345;
            foreach (Entry a in dirA)
            {
                String aName = a.Name;
                if (a.IsDirectoryEntry)
                {
                    aSizes.Add(aName, isDirectory);
                }
                else
                {
                    aSizes.Add(aName, ((DocumentNode)a).Size);
                }
            }
            foreach (Entry b in dirB)
            {
                String bName = b.Name;
                if (!aSizes.ContainsKey(bName))
                {
                    // In B but not A
                    return false;
                }

                int size;
                if (b.IsDirectoryEntry)
                {
                    size = isDirectory;
                }
                else
                {
                    size = ((DocumentNode)b).Size;
                }
                if (size != aSizes[(bName)])
                {
                    // Either the wrong type, or they're different sizes
                    return false;
                }

                // Track it as checked
                aSizes.Remove(bName);
            }
            if (!(aSizes.Count == 0))
            {
                // Nodes were in A but not B
                return false;
            }

            // If that passed, check entry contents
            foreach (Entry a in dirA)
            {
                try
                {
                    Entry b = dirB.GetEntry(a.Name);
                    bool match;
                    if (a.IsDirectoryEntry)
                    {
                        match = AreDirectoriesIdentical(
                              (DirectoryEntry)a, (DirectoryEntry)b);
                    }
                    else
                    {
                        match = AreDocumentsIdentical(
                              (DocumentEntry)a, (DocumentEntry)b);
                    }
                    if (!match) return false;
                }
                catch (FileNotFoundException)
                {
                    // Shouldn't really happen...
                    return false;
                }
                catch (IOException)
                {
                    // Something's messed up with one document, not a match
                    return false;
                }
            }

            // If we get here, they match!
            return true;
        }

        /**
         * Checks to see if two Documents have the same name
         *  and the same contents. (Their parent directories are
         *  not checked)
         */
        public static bool AreDocumentsIdentical(DocumentEntry docA, DocumentEntry docB)
        {
            if (!docA.Name.Equals(docB.Name))
            {
                // Names don't match, not the same
                return false;
            }
            if (docA.Size != docB.Size)
            {
                // Wrong sizes, can't have the same contents
                return false;
            }

            bool matches = true;
            DocumentInputStream inpA = null, inpB = null;
            try
            {
                inpA = new DocumentInputStream(docA);
                inpB = new DocumentInputStream(docB);

                int readA, readB;
                do
                {
                    readA = inpA.Read();
                    readB = inpB.Read();
                    if (readA != readB)
                    {
                        matches = false;
                        break;
                    }
                } while (readA != -1 && readB != -1);
            }
            finally
            {
                if (inpA != null) inpA.Close();
                if (inpB != null) inpB.Close();
            }

            return matches;
        }
    }
}
