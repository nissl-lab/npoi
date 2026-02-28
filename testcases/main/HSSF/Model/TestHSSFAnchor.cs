/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.DDF;
using NUnit.Framework;using NUnit.Framework.Legacy;
using TestCases.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace TestCases.HSSF.Model
{
    [TestFixture]
    public class TestHSSFAnchor
    {
        [Test]
        public void TestDefaultValues()
        {
            HSSFClientAnchor clientAnchor = new HSSFClientAnchor();
            ClassicAssert.AreEqual((int)clientAnchor.AnchorType, 0);
            ClassicAssert.AreEqual(clientAnchor.Col1, 0);
            ClassicAssert.AreEqual(clientAnchor.Col2, 0);
            ClassicAssert.AreEqual(clientAnchor.Dx1, 0);
            ClassicAssert.AreEqual(clientAnchor.Dx2, 0);
            ClassicAssert.AreEqual(clientAnchor.Dy1, 0);
            ClassicAssert.AreEqual(clientAnchor.Dy2, 0);
            ClassicAssert.AreEqual(clientAnchor.Row1, 0);
            ClassicAssert.AreEqual(clientAnchor.Row2, 0);

            clientAnchor = new HSSFClientAnchor(new EscherClientAnchorRecord());
            ClassicAssert.AreEqual((int)clientAnchor.AnchorType, 0);
            ClassicAssert.AreEqual(clientAnchor.Col1, 0);
            ClassicAssert.AreEqual(clientAnchor.Col2, 0);
            ClassicAssert.AreEqual(clientAnchor.Dx1, 0);
            ClassicAssert.AreEqual(clientAnchor.Dx2, 0);
            ClassicAssert.AreEqual(clientAnchor.Dy1, 0);
            ClassicAssert.AreEqual(clientAnchor.Dy2, 0);
            ClassicAssert.AreEqual(clientAnchor.Row1, 0);
            ClassicAssert.AreEqual(clientAnchor.Row2, 0);

            HSSFChildAnchor childAnchor = new HSSFChildAnchor();
            ClassicAssert.AreEqual(childAnchor.Dx1, 0);
            ClassicAssert.AreEqual(childAnchor.Dx2, 0);
            ClassicAssert.AreEqual(childAnchor.Dy1, 0);
            ClassicAssert.AreEqual(childAnchor.Dy2, 0);

            childAnchor = new HSSFChildAnchor(new EscherChildAnchorRecord());
            ClassicAssert.AreEqual(childAnchor.Dx1, 0);
            ClassicAssert.AreEqual(childAnchor.Dx2, 0);
            ClassicAssert.AreEqual(childAnchor.Dy1, 0);
            ClassicAssert.AreEqual(childAnchor.Dy2, 0);
        }
        [Test]
        public void TestCorrectOrderInSpContainer()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("drawings.xls");
            HSSFSheet sheet = wb.GetSheet("pictures") as HSSFSheet;
            HSSFPatriarch drawing = sheet.DrawingPatriarch as HSSFPatriarch;

            HSSFSimpleShape rectangle = (HSSFSimpleShape)drawing.Children[0];

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherContainer(rectangle).GetChild(0).RecordId, EscherSpRecord.RECORD_ID);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherContainer(rectangle).GetChild(1).RecordId, EscherOptRecord.RECORD_ID);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherContainer(rectangle).GetChild(2).RecordId, EscherClientAnchorRecord.RECORD_ID);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherContainer(rectangle).GetChild(3).RecordId, EscherClientDataRecord.RECORD_ID);

            rectangle.Anchor = (new HSSFClientAnchor());

            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherContainer(rectangle).GetChild(0).RecordId, EscherSpRecord.RECORD_ID);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherContainer(rectangle).GetChild(1).RecordId, EscherOptRecord.RECORD_ID);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherContainer(rectangle).GetChild(2).RecordId, EscherClientAnchorRecord.RECORD_ID);
            ClassicAssert.AreEqual(HSSFTestHelper.GetEscherContainer(rectangle).GetChild(3).RecordId, EscherClientDataRecord.RECORD_ID);
        }
        [Test]
        public void TestCreateClientAnchorFromContainer()
        {
            EscherContainerRecord container = new EscherContainerRecord();
            EscherClientAnchorRecord escher = new EscherClientAnchorRecord();
            escher.Flag=((short)3);
            escher.Col1=((short)11);
            escher.Col2=((short)12);
            escher.Row1=((short)13);
            escher.Row2=((short)14);
            escher.Dx1=((short)15);
            escher.Dx2=((short)16);
            escher.Dy1=((short)17);
            escher.Dy2=((short)18);
            container.AddChildRecord(escher);

            HSSFClientAnchor anchor = (HSSFClientAnchor)HSSFAnchor.CreateAnchorFromEscher(container);
            ClassicAssert.AreEqual(anchor.Col1, 11);
            ClassicAssert.AreEqual(escher.Col1, 11);
            ClassicAssert.AreEqual(anchor.Col2, 12);
            ClassicAssert.AreEqual(escher.Col2, 12);
            ClassicAssert.AreEqual(anchor.Row1, 13);
            ClassicAssert.AreEqual(escher.Row1, 13);
            ClassicAssert.AreEqual(anchor.Row2, 14);
            ClassicAssert.AreEqual(escher.Row2, 14);
            ClassicAssert.AreEqual(anchor.Dx1, 15);
            ClassicAssert.AreEqual(escher.Dx1, 15);
            ClassicAssert.AreEqual(anchor.Dx2, 16);
            ClassicAssert.AreEqual(escher.Dx2, 16);
            ClassicAssert.AreEqual(anchor.Dy1, 17);
            ClassicAssert.AreEqual(escher.Dy1, 17);
            ClassicAssert.AreEqual(anchor.Dy2, 18);
            ClassicAssert.AreEqual(escher.Dy2, 18);
        }
        [Test]
        public void TestCreateChildAnchorFromContainer()
        {
            EscherContainerRecord container = new EscherContainerRecord();
            EscherChildAnchorRecord escher = new EscherChildAnchorRecord();
            escher.Dx1=((short)15);
            escher.Dx2=((short)16);
            escher.Dy1=((short)17);
            escher.Dy2=((short)18);
            container.AddChildRecord(escher);

            HSSFChildAnchor anchor = (HSSFChildAnchor)HSSFAnchor.CreateAnchorFromEscher(container);
            ClassicAssert.AreEqual(anchor.Dx1, 15);
            ClassicAssert.AreEqual(escher.Dx1, 15);
            ClassicAssert.AreEqual(anchor.Dx2, 16);
            ClassicAssert.AreEqual(escher.Dx2, 16);
            ClassicAssert.AreEqual(anchor.Dy1, 17);
            ClassicAssert.AreEqual(escher.Dy1, 17);
            ClassicAssert.AreEqual(anchor.Dy2, 18);
            ClassicAssert.AreEqual(escher.Dy2, 18);
        }
        [Test]
        public void TestShapeEscherMustHaveAnchorRecord()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.CreateSheet() as HSSFSheet;

            HSSFPatriarch drawing = sheet.CreateDrawingPatriarch() as HSSFPatriarch;
            HSSFClientAnchor anchor = new HSSFClientAnchor(10, 10, 200, 200, (short)2, 2, (short)15, 15);
            anchor.AnchorType = (AnchorType)(2);

            HSSFSimpleShape rectangle = drawing.CreateSimpleShape(anchor);
            rectangle.ShapeType = (HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);

            rectangle.Anchor = (anchor);

            ClassicAssert.IsNotNull(HSSFTestHelper.GetEscherAnchor(anchor));
            ClassicAssert.IsNotNull(HSSFTestHelper.GetEscherContainer(rectangle));
            ClassicAssert.IsTrue(HSSFTestHelper.GetEscherAnchor(anchor).Equals(HSSFTestHelper.GetEscherContainer(rectangle).GetChildById(EscherClientAnchorRecord.RECORD_ID)));
        }
        [Test]
        public void TestClientAnchorFromEscher()
        {
            EscherClientAnchorRecord escher = new EscherClientAnchorRecord();
            escher.Col1=((short)11);
            escher.Col2=((short)12);
            escher.Row1=((short)13);
            escher.Row2=((short)14);
            escher.Dx1=((short)15);
            escher.Dx2=((short)16);
            escher.Dy1=((short)17);
            escher.Dy2=((short)18);

            HSSFClientAnchor anchor = new HSSFClientAnchor(escher);
            ClassicAssert.AreEqual(anchor.Col1, 11);
            ClassicAssert.AreEqual(escher.Col1, 11);
            ClassicAssert.AreEqual(anchor.Col2, 12);
            ClassicAssert.AreEqual(escher.Col2, 12);
            ClassicAssert.AreEqual(anchor.Row1, 13);
            ClassicAssert.AreEqual(escher.Row1, 13);
            ClassicAssert.AreEqual(anchor.Row2, 14);
            ClassicAssert.AreEqual(escher.Row2, 14);
            ClassicAssert.AreEqual(anchor.Dx1, 15);
            ClassicAssert.AreEqual(escher.Dx1, 15);
            ClassicAssert.AreEqual(anchor.Dx2, 16);
            ClassicAssert.AreEqual(escher.Dx2, 16);
            ClassicAssert.AreEqual(anchor.Dy1, 17);
            ClassicAssert.AreEqual(escher.Dy1, 17);
            ClassicAssert.AreEqual(anchor.Dy2, 18);
            ClassicAssert.AreEqual(escher.Dy2, 18);
        }
        [Test]
        public void TestClientAnchorFromScratch()
        {
            HSSFClientAnchor anchor = new HSSFClientAnchor();
            EscherClientAnchorRecord escher = (EscherClientAnchorRecord)HSSFTestHelper.GetEscherAnchor(anchor);
            anchor.SetAnchor((short)11, 12, 13, 14, (short)15, 16, 17, 18);

            ClassicAssert.AreEqual(anchor.Col1, 11);
            ClassicAssert.AreEqual(escher.Col1, 11);
            ClassicAssert.AreEqual(anchor.Col2, 15);
            ClassicAssert.AreEqual(escher.Col2, 15);
            ClassicAssert.AreEqual(anchor.Row1, 12);
            ClassicAssert.AreEqual(escher.Row1, 12);
            ClassicAssert.AreEqual(anchor.Row2, 16);
            ClassicAssert.AreEqual(escher.Row2, 16);
            ClassicAssert.AreEqual(anchor.Dx1, 13);
            ClassicAssert.AreEqual(escher.Dx1, 13);
            ClassicAssert.AreEqual(anchor.Dx2, 17);
            ClassicAssert.AreEqual(escher.Dx2, 17);
            ClassicAssert.AreEqual(anchor.Dy1, 14);
            ClassicAssert.AreEqual(escher.Dy1, 14);
            ClassicAssert.AreEqual(anchor.Dy2, 18);
            ClassicAssert.AreEqual(escher.Dy2, 18);

            anchor.Col1=(111);
            ClassicAssert.AreEqual(anchor.Col1, 111);
            ClassicAssert.AreEqual(escher.Col1, 111);
            anchor.Col2=(112);
            ClassicAssert.AreEqual(anchor.Col2, 112);
            ClassicAssert.AreEqual(escher.Col2, 112);
            anchor.Row1=(113);
            ClassicAssert.AreEqual(anchor.Row1, 113);
            ClassicAssert.AreEqual(escher.Row1, 113);
            anchor.Row2=(114);
            ClassicAssert.AreEqual(anchor.Row2, 114);
            ClassicAssert.AreEqual(escher.Row2, 114);
            anchor.Dx1=(115);
            ClassicAssert.AreEqual(anchor.Dx1, 115);
            ClassicAssert.AreEqual(escher.Dx1, 115);
            anchor.Dx2=(116);
            ClassicAssert.AreEqual(anchor.Dx2, 116);
            ClassicAssert.AreEqual(escher.Dx2, 116);
            anchor.Dy1=(117);
            ClassicAssert.AreEqual(anchor.Dy1, 117);
            ClassicAssert.AreEqual(escher.Dy1, 117);
            anchor.Dy2=(118);
            ClassicAssert.AreEqual(anchor.Dy2, 118);
            ClassicAssert.AreEqual(escher.Dy2, 118);
        }
        [Test]
        public void TestChildAnchorFromEscher()
        {
            EscherChildAnchorRecord escher = new EscherChildAnchorRecord();
            escher.Dx1=((short)15);
            escher.Dx2=((short)16);
            escher.Dy1=((short)17);
            escher.Dy2=((short)18);

            HSSFChildAnchor anchor = new HSSFChildAnchor(escher);
            ClassicAssert.AreEqual(anchor.Dx1, 15);
            ClassicAssert.AreEqual(escher.Dx1, 15);
            ClassicAssert.AreEqual(anchor.Dx2, 16);
            ClassicAssert.AreEqual(escher.Dx2, 16);
            ClassicAssert.AreEqual(anchor.Dy1, 17);
            ClassicAssert.AreEqual(escher.Dy1, 17);
            ClassicAssert.AreEqual(anchor.Dy2, 18);
            ClassicAssert.AreEqual(escher.Dy2, 18);
        }
        [Test]
        public void TestChildAnchorFromScratch()
        {
            HSSFChildAnchor anchor = new HSSFChildAnchor();
            EscherChildAnchorRecord escher = (EscherChildAnchorRecord)HSSFTestHelper.GetEscherAnchor(anchor);
            anchor.SetAnchor(11, 12, 13, 14);

            ClassicAssert.AreEqual(anchor.Dx1, 11);
            ClassicAssert.AreEqual(escher.Dx1, 11);
            ClassicAssert.AreEqual(anchor.Dx2, 13);
            ClassicAssert.AreEqual(escher.Dx2, 13);
            ClassicAssert.AreEqual(anchor.Dy1, 12);
            ClassicAssert.AreEqual(escher.Dy1, 12);
            ClassicAssert.AreEqual(anchor.Dy2, 14);
            ClassicAssert.AreEqual(escher.Dy2, 14);

            anchor.Dx1=(115);
            ClassicAssert.AreEqual(anchor.Dx1, 115);
            ClassicAssert.AreEqual(escher.Dx1, 115);
            anchor.Dx2=(116);
            ClassicAssert.AreEqual(anchor.Dx2, 116);
            ClassicAssert.AreEqual(escher.Dx2, 116);
            anchor.Dy1=(117);
            ClassicAssert.AreEqual(anchor.Dy1, 117);
            ClassicAssert.AreEqual(escher.Dy1, 117);
            anchor.Dy2=(118);
            ClassicAssert.AreEqual(anchor.Dy2, 118);
            ClassicAssert.AreEqual(escher.Dy2, 118);
        }
        [Test]
        public void TestEqualsToSelf()
        {
            HSSFClientAnchor clientAnchor = new HSSFClientAnchor(0, 1, 2, 3, (short)4, 5, (short)6, 7);
            ClassicAssert.AreEqual(clientAnchor, clientAnchor);

            HSSFChildAnchor childAnchor = new HSSFChildAnchor(0, 1, 2, 3);
            ClassicAssert.AreEqual(childAnchor, childAnchor);
        }
        [Test]
        public void TestPassIncompatibleTypeIsFalse()
        {
            HSSFClientAnchor clientAnchor = new HSSFClientAnchor(0, 1, 2, 3, (short)4, 5, (short)6, 7);
            ClassicAssert.AreNotSame(clientAnchor, "wrongType");

            HSSFChildAnchor childAnchor = new HSSFChildAnchor(0, 1, 2, 3);
            ClassicAssert.AreNotSame(childAnchor, "wrongType");
        }
        [Test]
        public void TestNullReferenceIsFalse()
        {
            HSSFClientAnchor clientAnchor = new HSSFClientAnchor(0, 1, 2, 3, (short)4, 5, (short)6, 7);
            ClassicAssert.IsFalse(clientAnchor.Equals(null), "Passing null to equals should return false");

            HSSFChildAnchor childAnchor = new HSSFChildAnchor(0, 1, 2, 3);
            ClassicAssert.IsFalse(childAnchor.Equals(null), "Passing null to equals should return false");
        }
        [Test]
        public void TestEqualsIsReflexiveIsSymmetric()
        {
            HSSFClientAnchor clientAnchor1 = new HSSFClientAnchor(0, 1, 2, 3, (short)4, 5, (short)6, 7);
            HSSFClientAnchor clientAnchor2 = new HSSFClientAnchor(0, 1, 2, 3, (short)4, 5, (short)6, 7);

            ClassicAssert.IsTrue(clientAnchor1.Equals(clientAnchor2));
            ClassicAssert.IsTrue(clientAnchor1.Equals(clientAnchor2));

            HSSFChildAnchor childAnchor1 = new HSSFChildAnchor(0, 1, 2, 3);
            HSSFChildAnchor childAnchor2 = new HSSFChildAnchor(0, 1, 2, 3);

            ClassicAssert.IsTrue(childAnchor1.Equals(childAnchor2));
            ClassicAssert.IsTrue(childAnchor2.Equals(childAnchor1));
        }
        [Test]
        public void TestEqualsValues()
        {
            HSSFClientAnchor clientAnchor1 = new HSSFClientAnchor(0, 1, 2, 3, (short)4, 5, (short)6, 7);
            HSSFClientAnchor clientAnchor2 = new HSSFClientAnchor(0, 1, 2, 3, (short)4, 5, (short)6, 7);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            clientAnchor2.Dx1=(10);
            ClassicAssert.AreNotSame(clientAnchor1, clientAnchor2);
            clientAnchor2.Dx1=(0);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            clientAnchor2.Dy1=(10);
            ClassicAssert.AreNotSame(clientAnchor1, clientAnchor2);
            clientAnchor2.Dy1=(1);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            clientAnchor2.Dx2=(10);
            ClassicAssert.AreNotSame(clientAnchor1, clientAnchor2);
            clientAnchor2.Dx2=(2);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            clientAnchor2.Dy2=(10);
            ClassicAssert.AreNotSame(clientAnchor1, clientAnchor2);
            clientAnchor2.Dy2=(3);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            clientAnchor2.Col1=(10);
            ClassicAssert.AreNotSame(clientAnchor1, clientAnchor2);
            clientAnchor2.Col1=(4);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            clientAnchor2.Row1=(10);
            ClassicAssert.AreNotSame(clientAnchor1, clientAnchor2);
            clientAnchor2.Row1=(5);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            clientAnchor2.Col2=(10);
            ClassicAssert.AreNotSame(clientAnchor1, clientAnchor2);
            clientAnchor2.Col2=(6);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            clientAnchor2.Row2=(10);
            ClassicAssert.AreNotSame(clientAnchor1, clientAnchor2);
            clientAnchor2.Row2=(7);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            clientAnchor2.AnchorType = (AnchorType)(3);
            ClassicAssert.AreNotSame(clientAnchor1, clientAnchor2);
            clientAnchor2.AnchorType=(0);
            ClassicAssert.AreEqual(clientAnchor1, clientAnchor2);

            HSSFChildAnchor childAnchor1 = new HSSFChildAnchor(0, 1, 2, 3);
            HSSFChildAnchor childAnchor2 = new HSSFChildAnchor(0, 1, 2, 3);

            childAnchor1.Dx1=(10);
            ClassicAssert.AreNotSame(childAnchor1, childAnchor2);
            childAnchor1.Dx1=(0);
            ClassicAssert.AreEqual(childAnchor1, childAnchor2);

            childAnchor2.Dy1=(10);
            ClassicAssert.AreNotSame(childAnchor1, childAnchor2);
            childAnchor2.Dy1=(1);
            ClassicAssert.AreEqual(childAnchor1, childAnchor2);

            childAnchor2.Dx2=(10);
            ClassicAssert.AreNotSame(childAnchor1, childAnchor2);
            childAnchor2.Dx2=(2);
            ClassicAssert.AreEqual(childAnchor1, childAnchor2);

            childAnchor2.Dy2=(10);
            ClassicAssert.AreNotSame(childAnchor1, childAnchor2);
            childAnchor2.Dy2=(3);
            ClassicAssert.AreEqual(childAnchor1, childAnchor2);
        }
        [Test]
        public void testFlipped()
        {
            HSSFChildAnchor child = new HSSFChildAnchor(2, 2, 1, 1);
            ClassicAssert.AreEqual(child.IsHorizontallyFlipped, true);
            ClassicAssert.AreEqual(child.IsVerticallyFlipped, true);
            ClassicAssert.AreEqual(child.Dx1, 1);
            ClassicAssert.AreEqual(child.Dx2, 2);
            ClassicAssert.AreEqual(child.Dy1, 1);
            ClassicAssert.AreEqual(child.Dy2, 2);

            child = new HSSFChildAnchor(3, 3, 4, 4);
            ClassicAssert.AreEqual(child.IsHorizontallyFlipped, false);
            ClassicAssert.AreEqual(child.IsVerticallyFlipped, false);
            ClassicAssert.AreEqual(child.Dx1, 3);
            ClassicAssert.AreEqual(child.Dx2, 4);
            ClassicAssert.AreEqual(child.Dy1, 3);
            ClassicAssert.AreEqual(child.Dy2, 4);

            HSSFClientAnchor client = new HSSFClientAnchor(1, 1, 1, 1, (short)4, 4, (short)3, 3);
            ClassicAssert.AreEqual(client.IsVerticallyFlipped, true);
            ClassicAssert.AreEqual(client.IsHorizontallyFlipped, true);
            ClassicAssert.AreEqual(client.Col1, 3);
            ClassicAssert.AreEqual(client.Col2, 4);
            ClassicAssert.AreEqual(client.Row1, 3);
            ClassicAssert.AreEqual(client.Row2, 4);

            client = new HSSFClientAnchor(1, 1, 1, 1, (short)5, 5, (short)6, 6);
            ClassicAssert.AreEqual(client.IsVerticallyFlipped, false);
            ClassicAssert.AreEqual(client.IsHorizontallyFlipped, false);
            ClassicAssert.AreEqual(client.Col1, 5);
            ClassicAssert.AreEqual(client.Col2, 6);
            ClassicAssert.AreEqual(client.Row1, 5);
            ClassicAssert.AreEqual(client.Row2, 6);
        }
    }
}
