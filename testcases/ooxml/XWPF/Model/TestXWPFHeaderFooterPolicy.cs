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

namespace NPOI.XWPF.Model
{
    using System;



using Microsoft.VisualStudio.TestTools.UnitTesting;

using NPOI.XWPF;
using NPOI.XWPF.UserModel;

/**
 * Tests for XWPF Header Footer Stuff
 */
    [TestClass]
    public class TestXWPFHeaderFooterPolicy 
{
	private XWPFDocument noHeader;
	private XWPFDocument header;
	private XWPFDocument headerFooter;
	private XWPFDocument footer;
	private XWPFDocument oddEven;
	private XWPFDocument diffFirst;

	protected void SetUp() throws IOException {

	    noHeader = XWPFTestDataSamples.OpenSampleDocument("NoHeadFoot.docx");
		header = XWPFTestDataSamples.OpenSampleDocument("ThreeColHead.docx");
		headerFooter = XWPFTestDataSamples.OpenSampleDocument("SimpleHeadThreeColFoot.docx");
		footer = XWPFTestDataSamples.OpenSampleDocument("FancyFoot.docx");
		oddEven = XWPFTestDataSamples.OpenSampleDocument("PageSpecificHeadFoot.docx");
		diffFirst = XWPFTestDataSamples.OpenSampleDocument("DiffFirstPageHeadFoot.docx");
	}

	    [TestMethod]
    public void TestPolicy(){
		XWPFHeaderFooterPolicy policy;

		policy = noHeader.HeaderFooterPolicy;
		Assert.IsNull(policy.DefaultHeader);
		Assert.IsNull(policy.DefaultFooter);

		Assert.IsNull(policy.GetHeader(1));
		Assert.IsNull(policy.GetHeader(2));
		Assert.IsNull(policy.GetHeader(3));
		Assert.IsNull(policy.GetFooter(1));
		Assert.IsNull(policy.GetFooter(2));
		Assert.IsNull(policy.GetFooter(3));


		policy = header.HeaderFooterPolicy;
		Assert.IsNotNull(policy.DefaultHeader);
		Assert.IsNull(policy.DefaultFooter);

		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(1));
		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(2));
		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(3));
		Assert.IsNull(policy.GetFooter(1));
		Assert.IsNull(policy.GetFooter(2));
		Assert.IsNull(policy.GetFooter(3));


		policy = footer.HeaderFooterPolicy;
		Assert.IsNull(policy.DefaultHeader);
		Assert.IsNotNull(policy.DefaultFooter);

		Assert.IsNull(policy.GetHeader(1));
		Assert.IsNull(policy.GetHeader(2));
		Assert.IsNull(policy.GetHeader(3));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(1));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(2));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(3));


		policy = headerFooter.HeaderFooterPolicy;
		Assert.IsNotNull(policy.DefaultHeader);
		Assert.IsNotNull(policy.DefaultFooter);

		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(1));
		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(2));
		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(3));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(1));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(2));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(3));


		policy = oddEven.HeaderFooterPolicy;
		Assert.IsNotNull(policy.DefaultHeader);
		Assert.IsNotNull(policy.DefaultFooter);
		Assert.IsNotNull(policy.EvenPageHeader);
		Assert.IsNotNull(policy.EvenPageFooter);

		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(1));
		Assert.AreEqual(policy.EvenPageHeader, policy.GetHeader(2));
		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(3));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(1));
		Assert.AreEqual(policy.EvenPageFooter, policy.GetFooter(2));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(3));


		policy = diffFirst.HeaderFooterPolicy;
		Assert.IsNotNull(policy.DefaultHeader);
		Assert.IsNotNull(policy.DefaultFooter);
		Assert.IsNotNull(policy.FirstPageHeader);
		Assert.IsNotNull(policy.FirstPageFooter);
		Assert.IsNull(policy.EvenPageHeader);
		Assert.IsNull(policy.EvenPageFooter);

		Assert.AreEqual(policy.FirstPageHeader, policy.GetHeader(1));
		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(2));
		Assert.AreEqual(policy.DefaultHeader, policy.GetHeader(3));
		Assert.AreEqual(policy.FirstPageFooter, policy.GetFooter(1));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(2));
		Assert.AreEqual(policy.DefaultFooter, policy.GetFooter(3));
	}

	    [TestMethod]
    public void TestContents(){
		XWPFHeaderFooterPolicy policy;

		// Test a few simple bits off a simple header
		policy = diffFirst.HeaderFooterPolicy;

		Assert.AreEqual(
			"I am the header on the first page, and I" + '\u2019' + "m nice and simple\n",
			policy.FirstPageHeader.Text
		);
		Assert.AreEqual(
				"First header column!\tMid header\tRight header!\n",
				policy.DefaultHeader.Text
		);


		// And a few bits off a more complex header
		policy = oddEven.HeaderFooterPolicy;

		Assert.AreEqual(
			"[ODD Page Header text]\n\n",
			policy.DefaultHeader.Text
		);
		Assert.AreEqual(
			"[This is an Even Page, with a Header]\n\n",
			policy.EvenPageHeader.Text
		);
	}
}

