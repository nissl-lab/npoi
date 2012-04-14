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

using NPOI.Openxml4j.exceptions;
using NPOI.Openxml4j.opc;
using NPOI.XSSF.UserModel;
using NPOI.XWPF;
using NPOI.XWPF.Model;

    [TestClass]
    public class TestXWPFPictureData 
{
    
    public void testRead() throws InvalidFormatException, IOException
    {
        XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("VariousPictures.docx");
        List<XWPFPictureData> pictures = sampleDoc.AllPictures;

        Assert.AreEqual(5,pictures.Size());
        String[] ext = {"wmf","png","emf","emf","jpeg"};
        for (int i = 0 ; i < pictures.Size() ; i++)
        {
            Assert.AreEqual(ext[i],pictures.Get(i).suggestFileExtension());
        }

        int num = pictures.Size();

        byte[] pictureData = XWPFTestDataSamples.GetImage("nature1.jpg");

        String relationId = sampleDoc.AddPictureData(pictureData,XWPFDocument.PICTURE_TYPE_JPEG);
        // picture list was updated
        Assert.AreEqual(num + 1,pictures.Size());
        XWPFPictureData pict = (XWPFPictureData) sampleDoc.GetRelationById(relationId);
        Assert.AreEqual("jpeg",pict.SuggestFileExtension());
        Assert.IsTrue(Arrays.Equals(pictureData,pict.Data));
    }

    public void testPictureInHeader() throws IOException
    {
        XWPFDocument sampleDoc = XWPFTestDataSamples.OpenSampleDocument("headerPic.docx");
        XWPFHeaderFooterPolicy policy = sampleDoc.HeaderFooterPolicy;

        XWPFHeader header = policy.DefaultHeader;

        List<XWPFPictureData> pictures = header.AllPictures;
        Assert.AreEqual(1,pictures.Size());
    }

    public void testNew() throws InvalidFormatException, IOException 
    {
        XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("EmptyDocumentWithHeaderFooter.docx");
        byte[] jpegData = XWPFTestDataSamples.GetImage("nature1.jpg");
        byte[] gifData = XWPFTestDataSamples.GetImage("nature1.gif");
        byte[] pngData = XWPFTestDataSamples.GetImage("nature1.png");

        List<XWPFPictureData> pictures = doc.AllPictures;
        Assert.AreEqual(0,pictures.Size());

        // Document shouldn't have any image relationships
        Assert.AreEqual(13,doc.PackagePart.Relationships.Size());
        foreach (PackageRelationship rel in doc.PackagePart.Relationships)
        {
            if (rel.RelationshipType.Equals(XSSFRelation.IMAGE_JPEG.Relation))
            {
                Assert.Fail("Shouldn't have JPEG yet");
            }
        }

        // Add the image
        String relationId = doc.AddPictureData(jpegData,XWPFDocument.PICTURE_TYPE_JPEG);
        Assert.AreEqual(1,pictures.Size());
        XWPFPictureData jpgPicData = (XWPFPictureData) doc.GetRelationById(relationId);
        Assert.AreEqual("jpeg",jpgPicData.SuggestFileExtension());
        Assert.IsTrue(Arrays.Equals(jpegData,jpgPicData.Data));

        // Ensure it now has one
        Assert.AreEqual(14,doc.PackagePart.Relationships.Size());
        PackageRelationship jpegRel = null;
        foreach (PackageRelationship rel in doc.PackagePart.Relationships)
        {
            if (rel.RelationshipType.Equals(XWPFRelation.IMAGE_JPEG.Relation))
            {
                if (jpegRel != null)
                    Assert.Fail("Found 2 jpegs!");
                jpegRel = rel;
            }
        }
        Assert.IsNotNull("JPEG Relationship not found",jpegRel);

        // Check the details
        Assert.AreEqual(XWPFRelation.IMAGE_JPEG.Relation,jpegRel.RelationshipType);
        Assert.AreEqual("/word/document.xml",jpegRel.Source.PartName.ToString());
        Assert.AreEqual("/word/media/image1.jpeg",jpegRel.TargetURI.Path);

        XWPFPictureData pictureDataByID = doc.GetPictureDataByID(jpegRel.Id);
        byte[] newJPEGData = pictureDataByID.Data;
        Assert.AreEqual(newJPEGData.Length,jpegData.Length);
        for (int i = 0 ; i < newJPEGData.Length ; i++)
        {
            Assert.AreEqual(newJPEGData[i],jpegData[i]);
        }

        // Save an re-load, check it appears
        doc = XWPFTestDataSamples.WriteOutAndReadBack(doc);
        Assert.AreEqual(1,doc.AllPictures.Size());
        Assert.AreEqual(1,doc.AllPackagePictures.Size());
    }
    
        [TestMethod]
    public void TestGetChecksum(){
        
    }
}

