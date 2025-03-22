/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is1 distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
namespace TestCases.HSSF.Model
{
    using System;
    using NPOI.DDF;
    using NPOI.HSSF.Model;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    [TestFixture]
    public class TestDrawingManager2
    {
        private DrawingManager2 drawingManager2;
        private EscherDggRecord dgg;
        [SetUp]
        public void SetUp()
        {

            dgg = new EscherDggRecord();
            dgg.FileIdClusters = (new EscherDggRecord.FileIdCluster[0]);
            drawingManager2 = new DrawingManager2(dgg);
        }
        [Test]
        public void TestCreateDgRecord()
        {
            EscherDgRecord dgRecord1 = drawingManager2.CreateDgRecord();
            ClassicAssert.AreEqual(1, dgRecord1.DrawingGroupId);
            ClassicAssert.AreEqual(-1, dgRecord1.LastMSOSPID);

            EscherDgRecord dgRecord2 = drawingManager2.CreateDgRecord();
            ClassicAssert.AreEqual(2, dgRecord2.DrawingGroupId);
            ClassicAssert.AreEqual(-1, dgRecord2.LastMSOSPID);

            ClassicAssert.AreEqual(2, dgg.DrawingsSaved);
            ClassicAssert.AreEqual(2, dgg.FileIdClusters.Length);
            ClassicAssert.AreEqual(3, dgg.NumIdClusters);
            ClassicAssert.AreEqual(0, dgg.NumShapesSaved);
        }
        [Test]
        public void TestAllocateShapeId()
        {
            EscherDgRecord dgRecord1 = drawingManager2.CreateDgRecord();
            EscherDgRecord dgRecord2 = drawingManager2.CreateDgRecord();

            ClassicAssert.AreEqual(1024, drawingManager2.AllocateShapeId((short)1));
            ClassicAssert.AreEqual(1024, dgRecord1.LastMSOSPID);
            ClassicAssert.AreEqual(1025, dgg.ShapeIdMax);
            ClassicAssert.AreEqual(1025, drawingManager2.AllocateShapeId((short)1));
            ClassicAssert.AreEqual(1025, dgRecord1.LastMSOSPID);
            ClassicAssert.AreEqual(1026, dgg.ShapeIdMax);
            ClassicAssert.AreEqual(1026, drawingManager2.AllocateShapeId((short)1));
            ClassicAssert.AreEqual(1026, dgRecord1.LastMSOSPID);
            ClassicAssert.AreEqual(1027, dgg.ShapeIdMax);
            ClassicAssert.AreEqual(2048, drawingManager2.AllocateShapeId((short)2));
            ClassicAssert.AreEqual(2048, dgRecord2.LastMSOSPID);
            ClassicAssert.AreEqual(2049, dgg.ShapeIdMax);

            for (int i = 0; i < 1021; i++)
            {
                drawingManager2.AllocateShapeId((short)1);
                ClassicAssert.AreEqual(2049, dgg.ShapeIdMax);
            }
            ClassicAssert.AreEqual(3072, drawingManager2.AllocateShapeId((short)1));
            ClassicAssert.AreEqual(3073, dgg.ShapeIdMax);

            ClassicAssert.AreEqual(2, dgg.DrawingsSaved);
            ClassicAssert.AreEqual(4, dgg.NumIdClusters);
            ClassicAssert.AreEqual(1026, dgg.NumShapesSaved);
        }
    }
}