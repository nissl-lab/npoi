/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
namespace TestCases.HSSF.Model
{
    using System;
    using NPOI.DDF;
    using NPOI.HSSF.Model;

    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NUnit.Framework.Legacy;

    [TestFixture]
    [Obsolete("deprecated in POI 3.15-beta2, scheduled for removal in 3.17, use DrawingManager2 instead")]
    public class TestDrawingManager
    {
        [Test]
        public void TestFindFreeSPIDBlock()
        {
            EscherDggRecord dgg = new EscherDggRecord();
            DrawingManager dm = new DrawingManager(dgg);
            dgg.ShapeIdMax = (1024);
            ClassicAssert.AreEqual(2048, dm.FindFreeSPIDBlock());
            dgg.ShapeIdMax = (1025);
            ClassicAssert.AreEqual(2048, dm.FindFreeSPIDBlock());
            dgg.ShapeIdMax = (2047);
            ClassicAssert.AreEqual(2048, dm.FindFreeSPIDBlock());
        }

        [Test]
        public void TestFindNewDrawingGroupId()
        {
            EscherDggRecord dgg = new EscherDggRecord();
            dgg.DrawingsSaved=(1);
            dgg.FileIdClusters=(new EscherDggRecord.FileIdCluster[]{
            new EscherDggRecord.FileIdCluster( 2, 10 )});
            DrawingManager dm = new DrawingManager(dgg);
            ClassicAssert.AreEqual(1, dm.FindNewDrawingGroupId());
            dgg.FileIdClusters=(new EscherDggRecord.FileIdCluster[]{
            new EscherDggRecord.FileIdCluster( 1, 10 ),
            new EscherDggRecord.FileIdCluster( 2, 10 )});
            ClassicAssert.AreEqual(3, dm.FindNewDrawingGroupId());
        }
        [Test]
        public void TestDrawingGroupExists()
        {
            EscherDggRecord dgg = new EscherDggRecord();
            dgg.DrawingsSaved=(1);
            dgg.FileIdClusters=(new EscherDggRecord.FileIdCluster[]{
            new EscherDggRecord.FileIdCluster( 2, 10 )});
            DrawingManager dm = new DrawingManager(dgg);
            ClassicAssert.IsFalse(dm.DrawingGroupExists((short)1));
            ClassicAssert.IsTrue(dm.DrawingGroupExists((short)2));
            ClassicAssert.IsFalse(dm.DrawingGroupExists((short)3));
        }
        [Test]
        public void TestCreateDgRecord()
        {
            EscherDggRecord dgg = new EscherDggRecord();
            dgg.DrawingsSaved=(0);
            dgg.FileIdClusters=(new EscherDggRecord.FileIdCluster[] { });
            DrawingManager dm = new DrawingManager(dgg);

            EscherDgRecord dgRecord = dm.CreateDgRecord();
            ClassicAssert.AreEqual(-1, dgRecord.LastMSOSPID);
            ClassicAssert.AreEqual(0, dgRecord.NumShapes);
            ClassicAssert.AreEqual(1, dm.Dgg.DrawingsSaved);
            ClassicAssert.AreEqual(1, dm.Dgg.FileIdClusters.Length);
            ClassicAssert.AreEqual(1, dm.Dgg.FileIdClusters[0].DrawingGroupId);
            ClassicAssert.AreEqual(0, dm.Dgg.FileIdClusters[0].NumShapeIdsUsed);
        }
        [Test]
        public void TestAllocateShapeId()
        {
            EscherDggRecord dgg = new EscherDggRecord();
            dgg.DrawingsSaved=(0);
            dgg.FileIdClusters=(new EscherDggRecord.FileIdCluster[] { });
            DrawingManager dm = new DrawingManager(dgg);

            EscherDgRecord dg = dm.CreateDgRecord();
            int shapeId = dm.AllocateShapeId(dg.DrawingGroupId);
            ClassicAssert.AreEqual(1024, shapeId);
            ClassicAssert.AreEqual(1025, dgg.ShapeIdMax);
            ClassicAssert.AreEqual(1, dgg.DrawingsSaved);
            ClassicAssert.AreEqual(1, dgg.FileIdClusters[0].DrawingGroupId);
            ClassicAssert.AreEqual(1, dgg.FileIdClusters[0].NumShapeIdsUsed);
            ClassicAssert.AreEqual(1024, dg.LastMSOSPID);
            ClassicAssert.AreEqual(1, dg.NumShapes);
        }

    }
}