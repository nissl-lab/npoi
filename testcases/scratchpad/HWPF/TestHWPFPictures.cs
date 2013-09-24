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

namespace TestCases.HWPF
{
using NPOI.HWPF.UserModel;

using System.Collections.Generic;
using NPOI.HWPF;
using NPOI.HWPF.Model;
    using System;
    using NUnit.Framework;

/**
 * Test picture support in HWPF
 * @author nick
 */
    [TestFixture]
public class TestHWPFPictures  {
	private String docAFile;
	private String docBFile;
	private String docCFile;
	private String docDFile;

	private String imgAFile;
	private String imgBFile;
	private String imgCFile;
	private String imgDFile;

        [SetUp]
	public void SetUp() {

		docAFile = "testPictures.doc";
		docBFile = "two_images.doc";
		docCFile = "vector_image.doc";
		docDFile = "GaiaTest.doc";

		imgAFile = "simple_image.jpg";
		imgBFile = "simple_image.png";
		imgCFile = "vector_image.emf";
		imgDFile = "GaiaTestImg.png";
	}

	/**
	 * Test just opening the files
	 */
    [Test]
	public void TestOpen() {
		HWPFTestDataSamples.OpenSampleFile(docAFile);
		HWPFTestDataSamples.OpenSampleFile(docBFile);
	}

	/**
	 * Test that we have the right numbers of images in each file
	 */
    [Test]
	public void TestImageCount() {
		HWPFDocument docA = HWPFTestDataSamples.OpenSampleFile(docAFile);
		HWPFDocument docB = HWPFTestDataSamples.OpenSampleFile(docBFile);

		Assert.IsNotNull(docA.GetPicturesTable());
		Assert.IsNotNull(docB.GetPicturesTable());

		PicturesTable picA = docA.GetPicturesTable();
		PicturesTable picB = docB.GetPicturesTable();

		List<Picture> picturesA = picA.GetAllPictures();
		List<Picture> picturesB = picB.GetAllPictures();

		Assert.AreEqual(7, picturesA.Count);
		Assert.AreEqual(2, picturesB.Count);
	}

	/**
	 * Test that we have the right images in at least one file
	 */
    [Test]
	public void TestImageData() {
		HWPFDocument docB = HWPFTestDataSamples.OpenSampleFile(docBFile);
		PicturesTable picB = docB.GetPicturesTable();
		List<Picture> picturesB = picB.GetAllPictures();

		Assert.AreEqual(2, picturesB.Count);

		Picture pic1 = picturesB[0];
		Picture pic2 = picturesB[1];

		Assert.IsNotNull(pic1);
		Assert.IsNotNull(pic2);

		// Check the same
		byte[] pic1B = ReadFile(imgAFile);
		byte[] pic2B = ReadFile(imgBFile);

		Assert.AreEqual(pic1B.Length, pic1.GetContent().Length);
		Assert.AreEqual(pic2B.Length, pic2.GetContent().Length);

		assertBytesSame(pic1B, pic1.GetContent());
		assertBytesSame(pic2B, pic2.GetContent());
	}

	/**
	 * Test that compressed image data is correctly returned.
	 */
    [Test]
	public void TestCompressedImageData() {
		HWPFDocument docC = HWPFTestDataSamples.OpenSampleFile(docCFile);
		PicturesTable picC = docC.GetPicturesTable();
		List<Picture> picturesC = picC.GetAllPictures();

		Assert.AreEqual(1, picturesC.Count);

		Picture pic = picturesC[0];
		Assert.IsNotNull(pic);

		// Check the same
		byte[] picBytes = ReadFile(imgCFile);

		Assert.AreEqual(picBytes.Length, pic.GetContent().Length);
		assertBytesSame(picBytes, pic.GetContent());
	}

	/**
	 * Pending the missing files being uploaded to
	 *  bug #44937
	 */
    //[Test]
	public void BROKENtestEscherDrawing() {
		HWPFDocument docD = HWPFTestDataSamples.OpenSampleFile(docDFile);
		List<Picture> allPictures = docD.GetPicturesTable().GetAllPictures();

		Assert.AreEqual(1, allPictures.Count);

		Picture pic = allPictures[0];
		Assert.IsNotNull(pic);
		byte[] picD = ReadFile(imgDFile);

		Assert.AreEqual(picD.Length, pic.GetContent().Length);

		assertBytesSame(picD, pic.GetContent());
	}

	private void assertBytesSame(byte[] a, byte[] b) {
		Assert.AreEqual(a.Length, b.Length);
		for(int i=0; i<a.Length; i++) {
			Assert.AreEqual(a[i],b[i]);
		}
	}

	private static byte[] ReadFile(String file) {
		return POIDataSamples.GetDocumentInstance().ReadFile(file);
	}
}

}
