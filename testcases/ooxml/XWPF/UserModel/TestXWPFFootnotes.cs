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

namespace NPOI.XWPF.UserModel
{
    using System;







using Microsoft.VisualStudio.TestTools.UnitTesting;

using NPOI.XWPF;

using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;

    [TestClass]
    public class TestXWPFFootnotes 
{

	    [TestMethod]
    public void TestAddFootnotesToDocument(){
		XWPFDocument docOut = new XWPFDocument();

		Bigint noteId = BigInt32.ValueOf(1);

		XWPFFootnotes footnotes = docOut.CreateFootnotes();
		CTFtnEdn ctNote = CTFtnEdn.Factory.NewInstance();
		ctNote.Id=(noteId);
		ctNote.Type=(STFtnEdn.NORMAL);
		footnotes.AddFootnote(ctNote);

		XWPFDocument docIn = XWPFTestDataSamples.WriteOutAndReadBack(docOut);

		XWPFFootnote note = docIn.GetFootnoteByID(noteId.IntValue());
		Assert.AreEqual(note.CTFtnEdn.Type, STFtnEdn.NORMAL);
	}
}


