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

using NPOI.HWPF;
using TestCases.HWPF;
using NPOI.HWPF.UserModel;
using NUnit.Framework;
namespace TestCases.HWPF.Model
{
    /**
     * Test cases for {@link NotesTables} and default implementation of
     * {@link Notes}
     * 
     * @author Sergey Vladimirov (vlsergey {at} gmail {dot} com)
     */
    [TestFixture]
    public class TestNotesTables
    {
        [Test]
        public void TestNotes()
        {
            HWPFDocument doc = HWPFTestDataSamples
                    .OpenSampleFile("endingnote.doc");
            Notes notes = doc.GetEndnotes();

            Assert.AreEqual(1, notes.GetNotesCount());

            Assert.AreEqual(10, notes.GetNoteAnchorPosition(0));
            Assert.AreEqual(0, notes.GetNoteTextStartOffSet(0));
            Assert.AreEqual(19, notes.GetNoteTextEndOffSet(0));
        }
    }
}

