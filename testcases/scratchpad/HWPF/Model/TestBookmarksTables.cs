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
using TestCases.HWPF;

using NPOI.HWPF.UserModel;
using NUnit.Framework;
namespace NPOI.HWPF.Model
{


    /**
     * Test cases for {@link BookmarksTables} and default implementation of
     * {@link Bookmarks}
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    [TestFixture]
    public class TestBookmarksTables 
    {
        [Test]
        public void TestBookmarks()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("pageref.doc");
            Bookmarks bookmarks = doc.GetBookmarks();

            Assert.AreEqual(1, bookmarks.Count);

            Bookmark bookmark = bookmarks.GetBookmark(0);
            Assert.AreEqual("userref", bookmark.Name);
            Assert.AreEqual(27, bookmark.Start);
            Assert.AreEqual(38, bookmark.End);
        }
    }


}