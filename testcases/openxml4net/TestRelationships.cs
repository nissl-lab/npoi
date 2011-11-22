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

using NPOI.Util;
using NPOI.OpenXml4Net.OPC;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestCases.OpenXml4Net;
using System.IO;
namespace TestCases.OPC
{


    [TestClass]
    public class TestRelationships  {
	private static String HYPERLINK_REL_TYPE =
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
	private static String COMMENTS_REL_TYPE =
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
	private static String SHEET_WITH_COMMENTS =
		"/xl/worksheets/sheet1.xml";

    private static POILogger logger = POILogFactory.GetLogger(typeof(TestPackageCoreProperties));

        /**
         * Test relationships are correctly loaded. This at the moment fails (as of r499)
         * whenever a document is loaded before its correspondig .rels file has been found.
         * The code in this case assumes there are no relationships defined, but it should
         * really look also for not yet loaded parts.
         */
        [TestMethod]
        public void TestLoadRelationships() {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("sample.xlsx");
            OPCPackage pkg = OPCPackage.Open(is1);
            logger.Log(POILogger.DEBUG, "1: " + pkg);
            PackageRelationshipCollection rels = pkg.GetRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT);
            PackageRelationship coreDocRelationship = rels.GetRelationship(0);
            PackagePart corePart = pkg.GetPart(coreDocRelationship);
            String[] relIds = { "rId1", "rId2", "rId3" };
            foreach (String relId in relIds) {
                PackageRelationship rel = corePart.GetRelationship(relId);
                Assert.IsNotNull(rel);
                PackagePartName relName = PackagingURIHelper.CreatePartName(rel.TargetUri);
                PackagePart sheetPart = pkg.GetPart(relName);
                Assert.AreEqual(1, sheetPart.Relationships.Size, "Number of relationships1 for " + sheetPart.PartName);
            }
        }
        
        /**
         * Checks that we can fetch a collection of relations by
         *  type, then grab from within there by id
         */
        [TestMethod]
        public void TestFetchFromCollection() {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("ExcelWithHyperlinks.xlsx");
            OPCPackage pkg = OPCPackage.Open(is1);
            PackagePart sheet = pkg.GetPart(
        		    PackagingURIHelper.CreatePartName(SHEET_WITH_COMMENTS));
            Assert.IsNotNull(sheet);
            
            Assert.IsTrue(sheet.HasRelationships);
            Assert.AreEqual(6, sheet.Relationships.Size);
            
            // Should have three hyperlinks, and one comment
            PackageRelationshipCollection hyperlinks =
        	    sheet.GetRelationshipsByType(HYPERLINK_REL_TYPE);
            PackageRelationshipCollection comments =
        	    sheet.GetRelationshipsByType(COMMENTS_REL_TYPE);
            Assert.AreEqual(3, hyperlinks.Size);
            Assert.AreEqual(1, comments.Size);
            
            // Check we can Get bits out by id
            // Hyperlinks are rId1, rId2 and rId3
            // Comment is rId6
            Assert.IsNotNull(hyperlinks.GetRelationshipByID("rId1"));
            Assert.IsNotNull(hyperlinks.GetRelationshipByID("rId2"));
            Assert.IsNotNull(hyperlinks.GetRelationshipByID("rId3"));
            Assert.IsNull(hyperlinks.GetRelationshipByID("rId6"));
            
            Assert.IsNull(comments.GetRelationshipByID("rId1"));
            Assert.IsNull(comments.GetRelationshipByID("rId2"));
            Assert.IsNull(comments.GetRelationshipByID("rId3"));
            Assert.IsNotNull(comments.GetRelationshipByID("rId6"));
            
            Assert.IsNotNull(sheet.GetRelationship("rId1"));
            Assert.IsNotNull(sheet.GetRelationship("rId2"));
            Assert.IsNotNull(sheet.GetRelationship("rId3"));
            Assert.IsNotNull(sheet.GetRelationship("rId6"));
        }
        
        /**
         * Excel uses relations on sheets to store the details of 
         *  external hyperlinks. Check we can load these ok.
         */
        [TestMethod]
        public void TestLoadExcelHyperlinkRelations() {
            Stream is1 = OpenXml4NetTestDataSamples.OpenSampleStream("ExcelWithHyperlinks.xlsx");
            OPCPackage pkg = OPCPackage.Open(is1);
	        PackagePart sheet = pkg.GetPart(
	    		    PackagingURIHelper.CreatePartName(SHEET_WITH_COMMENTS));
	        Assert.IsNotNull(sheet);

	        // rId1 is url
	        PackageRelationship url = sheet.GetRelationship("rId1");
	        Assert.IsNotNull(url);
	        Assert.AreEqual("rId1", url.Id);
	        Assert.AreEqual("/xl/worksheets/sheet1.xml", url.SourceUri.ToString());
	        Assert.AreEqual("http://poi.apache.org/", url.TargetUri.ToString());
    	    
	        // rId2 is file
	        PackageRelationship file = sheet.GetRelationship("rId2");
	        Assert.IsNotNull(file);
	        Assert.AreEqual("rId2", file.Id);
	        Assert.AreEqual("/xl/worksheets/sheet1.xml", file.SourceUri.ToString());
	        Assert.AreEqual("WithVariousData.xlsx", file.TargetUri.ToString());
    	    
	        // rId3 is mailto
	        PackageRelationship mailto = sheet.GetRelationship("rId3");
	        Assert.IsNotNull(mailto);
	        Assert.AreEqual("rId3", mailto.Id);
	        Assert.AreEqual("/xl/worksheets/sheet1.xml", mailto.SourceUri.ToString());
	        Assert.AreEqual("mailto:dev@poi.apache.org?subject=XSSF%20Hyperlinks", mailto.TargetUri.ToString());
        }
    
        /*
         * Excel uses relations on sheets to store the details of 
         *  external hyperlinks. Check we can create these OK, 
         *  then still read them later
         */
        [TestMethod]
        public void TestCreateExcelHyperlinkRelations() {
    	    String filepath = OpenXml4NetTestDataSamples.GetSampleFileName("ExcelWithHyperlinks.xlsx");
	        OPCPackage pkg = OPCPackage.Open(filepath, PackageAccess.READ_WRITE);
	        PackagePart sheet = pkg.GetPart(
	    		    PackagingURIHelper.CreatePartName(SHEET_WITH_COMMENTS));
	        Assert.IsNotNull(sheet);
    	    
	        Assert.AreEqual(3, sheet.GetRelationshipsByType(HYPERLINK_REL_TYPE).Size);
    	    
	        // Add three new ones
	        PackageRelationship openxml4j =
	    	    sheet.AddExternalRelationship("http://www.Openxml4j.org/", HYPERLINK_REL_TYPE);
	        PackageRelationship sf =
	    	    sheet.AddExternalRelationship("http://openxml4j.sf.net/", HYPERLINK_REL_TYPE);
	        PackageRelationship file =
	    	    sheet.AddExternalRelationship("MyDocument.docx", HYPERLINK_REL_TYPE);
    	    
	        // Check they were Added properly
	        Assert.IsNotNull(openxml4j);
	        Assert.IsNotNull(sf);
	        Assert.IsNotNull(file);
    	    
	        Assert.AreEqual(6, sheet.GetRelationshipsByType(HYPERLINK_REL_TYPE).Size);
    	    
	        Assert.AreEqual("http://www.openxml4j.org/", openxml4j.TargetUri.ToString());
	        Assert.AreEqual("/xl/worksheets/sheet1.xml", openxml4j.SourceUri.ToString());
	        Assert.AreEqual(HYPERLINK_REL_TYPE, openxml4j.RelationshipType);
    	    
	        Assert.AreEqual("http://openxml4j.sf.net/", sf.TargetUri.ToString());
	        Assert.AreEqual("/xl/worksheets/sheet1.xml", sf.SourceUri.ToString());
	        Assert.AreEqual(HYPERLINK_REL_TYPE, sf.RelationshipType);
    	    
	        Assert.AreEqual("MyDocument.docx", file.TargetUri.ToString());
	        Assert.AreEqual("/xl/worksheets/sheet1.xml", file.SourceUri.ToString());
	        Assert.AreEqual(HYPERLINK_REL_TYPE, file.RelationshipType);
    	    
	        // Will Get ids 7, 8 and 9, as we already have 1-6
	        Assert.AreEqual("rId7", openxml4j.Id);
	        Assert.AreEqual("rId8", sf.Id);
	        Assert.AreEqual("rId9", file.Id);
    	    
    	    
	        // Write out and re-load
	        MemoryStream baos = new MemoryStream();
	        pkg.Save(baos);
	        MemoryStream bais = new MemoryStream(baos.ToArray());
	        pkg = OPCPackage.Open(bais);
    	    
	        // Check again
	        sheet = pkg.GetPart(
	    		    PackagingURIHelper.CreatePartName(SHEET_WITH_COMMENTS));
    	    
	        Assert.AreEqual(6, sheet.GetRelationshipsByType(HYPERLINK_REL_TYPE).Size);
    	    
	        Assert.AreEqual("http://poi.apache.org/",
	    		    sheet.GetRelationship("rId1").TargetUri.ToString());
	        Assert.AreEqual("mailto:dev@poi.apache.org?subject=XSSF%20Hyperlinks",
	    		    sheet.GetRelationship("rId3").TargetUri.ToString());
    	    
	        Assert.AreEqual("http://www.Openxml4j.org/",
	    		    sheet.GetRelationship("rId7").TargetUri.ToString());
	        Assert.AreEqual("http://openxml4j.sf.net/",
	    		    sheet.GetRelationship("rId8").TargetUri.ToString());
	        Assert.AreEqual("MyDocument.docx",
	    		    sheet.GetRelationship("rId9").TargetUri.ToString());
        }
        [TestMethod]
        public void TestCreateRelationsFromScratch() {
    	    MemoryStream baos = new MemoryStream();
    	    OPCPackage pkg = OPCPackage.Create(baos);
        	
    	    PackagePart partA =
    		    pkg.CreatePart(PackagingURIHelper.CreatePartName("/partA"), "text/plain");
    	    PackagePart partB =
    		    pkg.CreatePart(PackagingURIHelper.CreatePartName("/partB"), "image/png");
    	    Assert.IsNotNull(partA);
    	    Assert.IsNotNull(partB);
        	
    	    // Internal
    	    partA.AddRelationship(partB.PartName, TargetMode.INTERNAL, "http://example/Rel");
        	
    	    // External
    	    partA.AddExternalRelationship("http://poi.apache.org/", "http://example/poi");
    	    partB.AddExternalRelationship("http://poi.apache.org/ss/", "http://example/poi/ss");

    	    // Check as expected currently
    	    Assert.AreEqual("/partB", partA.GetRelationship("rId1").TargetUri.ToString());
    	    Assert.AreEqual("http://poi.apache.org/", 
    			    partA.GetRelationship("rId2").TargetUri.ToString());
    	    Assert.AreEqual("http://poi.apache.org/ss/", 
    			    partB.GetRelationship("rId1").TargetUri.ToString());
        	
        	
    	    // Save, and re-load
    	    pkg.Close();
    	    MemoryStream bais = new MemoryStream(baos.ToArray());
    	    pkg = OPCPackage.Open(bais);
        	
    	    partA = pkg.GetPart(PackagingURIHelper.CreatePartName("/partA"));
    	    partB = pkg.GetPart(PackagingURIHelper.CreatePartName("/partB"));
        	
        	
    	    // Check the relations
    	    Assert.AreEqual(2, partA.Relationships.Size);
    	    Assert.AreEqual(1, partB.Relationships.Size);
        	
    	    Assert.AreEqual("/partB", partA.GetRelationship("rId1").TargetUri.ToString());
    	    Assert.AreEqual("http://poi.apache.org/", 
    			    partA.GetRelationship("rId2").TargetUri.ToString());
    	    Assert.AreEqual("http://poi.apache.org/ss/", 
    			    partB.GetRelationship("rId1").TargetUri.ToString());
    	    // Check core too
    	    Assert.AreEqual("/docProps/core.xml",
    			    pkg.GetRelationshipsByType("http://schemas.Openxmlformats.org/namespace/2006/relationships/metadata/core-properties").GetRelationship(0).TargetUri.ToString());
        }

        [TestMethod]
        public void TestTargetWithSpecialChars(){

            OPCPackage pkg;

            String filepath = OpenXml4NetTestDataSamples.GetSampleFileName("50154.xlsx");
            pkg = OPCPackage.Open(filepath);
            Assert_50154(pkg);

            MemoryStream baos = new MemoryStream();
            pkg.Save(baos);
            MemoryStream bais = new MemoryStream(baos.ToArray());
            pkg = OPCPackage.Open(bais);

            Assert_50154(pkg);
        }

        public void Assert_50154(OPCPackage pkg) {
            Uri drawingUri = new Uri("/xl/drawings/drawing1.xml",UriKind.Relative);
            PackagePart drawingPart = pkg.GetPart(PackagingURIHelper.CreatePartName(drawingUri));
            PackageRelationshipCollection drawingRels = drawingPart.Relationships;

            Assert.AreEqual(6, drawingRels.Size);

            // expected one image
            Assert.AreEqual(1, drawingPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image").Size);
            // and three hyperlinks
            Assert.AreEqual(5, drawingPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink").Size);

            PackageRelationship rId1 = drawingPart.GetRelationship("rId1");
            Uri parent = drawingPart.PartName.URI;
            Uri rel1 = new Uri(parent,rId1.TargetUri);
            Uri rel11 = PackagingURIHelper.RelativizeUri(drawingPart.PartName.URI, rId1.TargetUri);
            Assert.AreEqual("'Another Sheet'!A1", rel1.Fragment);

            PackageRelationship rId2 = drawingPart.GetRelationship("rId2");
            Uri rel2 = PackagingURIHelper.RelativizeUri(drawingPart.PartName.URI, rId2.TargetUri);
            Assert.AreEqual("../media/image1.png", rel2.OriginalString);

            PackageRelationship rId3 = drawingPart.GetRelationship("rId3");
            Uri rel3 = new Uri(parent,rId3.TargetUri);
            Assert.AreEqual("ThirdSheet!A1", rel3.Fragment);

            PackageRelationship rId4 = drawingPart.GetRelationship("rId4");
            Uri rel4 = new Uri(parent,rId4.TargetUri);
            Assert.AreEqual("'\u0410\u043F\u0430\u0447\u0435 \u041F\u041E\u0418'!A1", rel4.Fragment);

            PackageRelationship rId5 = drawingPart.GetRelationship("rId5");
            Uri rel5 = new Uri(parent,rId5.TargetUri);
            // back slashed have been Replaced with forward
            Assert.AreEqual("file:///D:/chan-chan.mp3", rel5.ToString());

            PackageRelationship rId6 = drawingPart.GetRelationship("rId6");
            Uri rel6 = new Uri(parent,rId6.TargetUri);
            Assert.AreEqual("../../../../../../../cygwin/home/yegor/dinom/&&&[access].2010-10-26.log", rel6.OriginalString);
            Assert.AreEqual("'\u0410\u043F\u0430\u0447\u0435 \u041F\u041E\u0418'!A5", rel6.Fragment);
        }
        [TestMethod]
        public void TestSelfRelations_bug51187() {
    	MemoryStream baos = new MemoryStream();
    	OPCPackage pkg = OPCPackage.Create(baos);

    	PackagePart partA =
    		pkg.CreatePart(PackagingURIHelper.CreatePartName("/partA"), "text/plain");
    	Assert.IsNotNull(partA);

    	// reference itself
    	PackageRelationship rel1 = partA.AddRelationship(partA.PartName, TargetMode.INTERNAL, "partA");

    	
    	// Save, and re-load
    	pkg.Close();
    	MemoryStream bais = new MemoryStream(baos.ToArray());
    	pkg = OPCPackage.Open(bais);

    	partA = pkg.GetPart(PackagingURIHelper.CreatePartName("/partA"));


    	// Check the relations
    	Assert.AreEqual(1, partA.Relationships.Size);

       PackageRelationship rel2 = partA.Relationships.GetRelationship(0);

    	Assert.AreEqual(rel1.RelationshipType, rel2.RelationshipType);
       Assert.AreEqual(rel1.Id, rel2.Id);
       Assert.AreEqual(rel1.SourceUri, rel2.SourceUri);
       Assert.AreEqual(rel1.TargetUri, rel2.TargetUri);
       Assert.AreEqual(rel1.TargetMode, rel2.TargetMode);
        }
    }
}

