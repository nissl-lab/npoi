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
using System;
using NPOI.HWPF.Model;
using System.Collections.Generic;
using NPOI.HWPF.UserModel;
using NPOI.Util;
namespace NPOI.HWPF.UserModel
{

    /**
     * Implementation of user-friendly interface for document bookmarks
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {doc} com)
     */
    public class BookmarksImpl : Bookmarks
    {

        private BookmarksTables bookmarksTables;

        private Dictionary<int, List<GenericPropertyNode>> sortedDescriptors = null;

        private int[] sortedStartPositions = null;

        public BookmarksImpl(BookmarksTables bookmarksTables)
        {
            this.bookmarksTables = bookmarksTables;
        }

        private Bookmark GetBookmark(GenericPropertyNode first)
        {
            return new BookmarkImpl(this.bookmarksTables, first);
        }

        public Bookmark GetBookmark(int index)
        {
            GenericPropertyNode first = bookmarksTables
                    .GetDescriptorFirst(index);
            return GetBookmark(first);
        }

        public List<Bookmark> GetBookmarksAt(int startCp)
        {
            UpdateSortedDescriptors();

            List<GenericPropertyNode> nodes = sortedDescriptors[startCp];
            if (nodes == null || nodes.Count == 0)
                return new List<Bookmark>();

            List<Bookmark> result = new List<Bookmark>(nodes.Count);
            foreach (GenericPropertyNode node in nodes)
            {
                result.Add(GetBookmark(node));
            }
            return result;
        }

        public int Count
        {
            get
            {
                return bookmarksTables.GetDescriptorsFirstCount();
            }
        }

        public Dictionary<int, List<Bookmark>> GetBookmarksStartedBetween(
                int startInclusive, int endExclusive)
        {
            UpdateSortedDescriptors();

            int startLookupIndex = Array.BinarySearch(this.sortedStartPositions,
                    startInclusive);
            if (startLookupIndex < 0)
                startLookupIndex = -(startLookupIndex + 1);
            int endLookupIndex = Array.BinarySearch(this.sortedStartPositions,
                    endExclusive);
            if (endLookupIndex < 0)
                endLookupIndex = -(endLookupIndex + 1);

            Dictionary<int, List<Bookmark>> result = new Dictionary<int, List<Bookmark>>();
            for (int LookupIndex = startLookupIndex; LookupIndex < endLookupIndex; LookupIndex++)
            {
                int s = sortedStartPositions[LookupIndex];
                if (s < startInclusive)
                    continue;
                if (s >= endExclusive)
                    break;

                List<Bookmark> startedAt = GetBookmarksAt(s);
                if (startedAt != null)
                    result[s] = startedAt;
            }

            return result;
        }

        private void UpdateSortedDescriptors()
        {
            if (sortedDescriptors != null)
                return;

            Dictionary<int, List<GenericPropertyNode>> result = new Dictionary<int, List<GenericPropertyNode>>();
            for (int b = 0; b < bookmarksTables.GetDescriptorsFirstCount(); b++)
            {
                GenericPropertyNode property = bookmarksTables
                        .GetDescriptorFirst(b);
                int positionKey = property.Start;
                List<GenericPropertyNode> atPositionList = result[positionKey];
                if (atPositionList == null)
                {
                    atPositionList = new List<GenericPropertyNode>();
                    result[positionKey] = atPositionList;
                }
                atPositionList.Add(property);
            }

            int counter = 0;
            int[] indices = new int[result.Count];
            foreach (KeyValuePair<int, List<GenericPropertyNode>> entry in result)
            {
                indices[counter++] = entry.Key;
                List<GenericPropertyNode> updated = new List<GenericPropertyNode>(
                        entry.Value);
                updated.Sort((IComparer<GenericPropertyNode>)PropertyNode.EndComparator.instance);
                result[entry.Key] = updated;
            }
            Array.Sort(indices);

            this.sortedDescriptors = result;
            this.sortedStartPositions = indices;
        }

        internal class BookmarkImpl : Bookmark
        {
            private GenericPropertyNode first;
            private BookmarksTables bookmarksTables;

            internal BookmarkImpl(BookmarksTables bookmarksTables, GenericPropertyNode first)
            {
                this.bookmarksTables = bookmarksTables;
                this.first = first;
            }

            public override bool Equals(Object obj)
            {
                if (this == obj)
                    return true;
                if (obj == null)
                    return false;
                if (this.GetType() != obj.GetType())
                    return false;
                BookmarkImpl other = (BookmarkImpl)obj;
                if (first == null)
                {
                    if (other.first != null)
                        return false;
                }
                else if (!first.Equals(other.first))
                    return false;
                return true;
            }

            public int End
            {
                get
                {
                    int currentIndex = bookmarksTables.GetDescriptorFirstIndex(first);
                    try
                    {
                        GenericPropertyNode descriptorLim = bookmarksTables
                                .GetDescriptorLim(currentIndex);
                        return descriptorLim.Start;
                    }
                    catch (IndexOutOfRangeException)
                    {
                        return first.End;
                    }
                }
            }

            public String Name
            {
                get
                {
                    int currentIndex = bookmarksTables.GetDescriptorFirstIndex(first);
                    try
                    {
                        return bookmarksTables.GetName(currentIndex);
                    }
                    catch (IndexOutOfRangeException)
                    {
                        return "";
                    }
                }
                set 
                {
                    int currentIndex = bookmarksTables.GetDescriptorFirstIndex(first);
                    bookmarksTables.SetName(currentIndex, value);   
                }
            }

            public int Start
            {
                get
                {
                    return first.Start;
                }
            }

            public override int GetHashCode()
            {
                return 31 + (first == null ? 0 : first.GetHashCode());
            }
            public override String ToString()
            {
                return "Bookmark [" + Start + "; " + End + "): name: "
                        + Name;
            }

        }
    }
}

