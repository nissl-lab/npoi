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
using System.Collections.Generic;
namespace NPOI.HWPF.UserModel
{
    /**
     * User-friendly interface to access document bookmarks
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    public interface Bookmarks
    {
        /**
         * @param index
         *            bookmark document index
         * @return {@link Bookmark} with specified index
         * @throws IndexOutOfBoundsException
         *             if bookmark with specified index not present in document
         */
        Bookmark GetBookmark(int index);

        /**
         * @return count of {@link Bookmark}s in document
         */
        int Count { get; }

        /**
         * @return {@link Map} of bookmarks started in specified range, where key is
         *         start position and value is sorted {@link List} of
         *         {@link Bookmark}
         */
        Dictionary<int, List<Bookmark>> GetBookmarksStartedBetween(
                int startInclusive, int endExclusive);
    }


}