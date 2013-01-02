/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

namespace TestCases.HSSF.Model
{

    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;

    /**
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestShapes
    {

        /**
         * Test generator of ids for the CommonObjectDataSubRecord record.
         *
         * See Bug 51332
         */
        [Test]
        public void TestShapeId()
        {

            HSSFClientAnchor anchor = new HSSFClientAnchor();
            AbstractShape shape;
            CommonObjectDataSubRecord cmo;

            shape = new TextboxShape(new HSSFTextbox(null, anchor), 1025);
            cmo = (CommonObjectDataSubRecord)shape.ObjRecord.SubRecords[(0)];
            Assert.AreEqual(1, cmo.ObjectId);

            shape = new PictureShape(new HSSFPicture(null, anchor), 1026);
            cmo = (CommonObjectDataSubRecord)shape.ObjRecord.SubRecords[(0)];
            Assert.AreEqual(2, cmo.ObjectId);

            shape = new CommentShape(new HSSFComment(null, anchor), 1027);
            cmo = (CommonObjectDataSubRecord)shape.ObjRecord.SubRecords[(0)];
            Assert.AreEqual(1027, cmo.ObjectId);
        }
    }

}