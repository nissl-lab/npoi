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
using System.Collections;

using NPOI.HWPF.UserModel;
using System.IO;
using NUnit.Framework;
namespace TestCases.HWPF.UserModel
{

    /**
     * Test the shapes handling
     */
    [TestFixture]
    public class TestShapes
    {

        /**
         * two shapes, second is a group
         */
        [Test]
        public void TestShapes1()
        {
            HWPFDocument doc = HWPFTestDataSamples.OpenSampleFile("WithArtShapes.doc");

            IList shapes = doc.GetShapesTable().GetAllShapes();
            IList vshapes = doc.GetShapesTable().GetVisibleShapes();

            Assert.AreEqual(2, shapes.Count);
            Assert.AreEqual(2, vshapes.Count);

            Shape s1 = (Shape)shapes[0];
            Shape s2 = (Shape)shapes[1];

            Assert.AreEqual(3616, s1.Width);
            Assert.AreEqual(1738, s1.Height);
            Assert.AreEqual(true, s1.IsWithinDocument);

            Assert.AreEqual(4817, s2.Width);
            Assert.AreEqual(2164, s2.Height);
            Assert.AreEqual(true, s2.IsWithinDocument);


            // Re-serialisze, check still there
            MemoryStream baos = new MemoryStream();
            doc.Write(baos);
            MemoryStream bais = new MemoryStream(baos.ToArray());
            doc = new HWPFDocument(bais);

            shapes = doc.GetShapesTable().GetAllShapes();
            vshapes = doc.GetShapesTable().GetVisibleShapes();

            Assert.AreEqual(2, shapes.Count);
            Assert.AreEqual(2, vshapes.Count);

            s1 = (Shape)shapes[0];
            s2 = (Shape)shapes[1];

            Assert.AreEqual(3616, s1.Width);
            Assert.AreEqual(1738, s1.Height);
            Assert.AreEqual(true, s1.IsWithinDocument);

            Assert.AreEqual(4817, s2.Width);
            Assert.AreEqual(2164, s2.Height);
            Assert.AreEqual(true, s2.IsWithinDocument);
        }
    }
}

