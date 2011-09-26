/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

namespace TestCases.HSSF.EventModel
{
    using System;
    using System.IO;
    using System.Collections;

    using NPOI.HSSF;
    using NPOI.HSSF.EventModel;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.UserModel;


    /**
     * Tests the ModelFactory.  
     * 
     * @author Andrew C. Oliver acoliver@apache.org
     */
    [TestClass]
    public class TestModelFactory
    {
        private ModelFactory factory;
        private HSSFWorkbook book;
        private Stream in1;
        private ArrayList models;

        /**
         * Tests that the listeners collection is Created
         * @param arg0
         */
        public TestModelFactory()
        {

            ModelFactory mf = new ModelFactory();
            Assert.IsTrue(mf.listeners != null, "listeners member cannot be null");
            Assert.IsTrue(mf.listeners is IList, "listeners member must be a List");
        }

        
        [TestInitialize]
        public void SetUp()
        {

            models = new ArrayList(3);
            factory = new ModelFactory();
            book = new HSSFWorkbook();
            MemoryStream stream = (MemoryStream)SetupRunFile(book);
            POIFSFileSystem fs = new POIFSFileSystem(
                                       new MemoryStream(stream.ToArray())
                                       );
            in1 = fs.CreatePOIFSDocumentReader("Workbook");
        }
        [TestCleanup]
        public void TearDown()
        {

            factory = null;
            book = null;
            in1 = null;
        }

        /**
         * Tests that listeners can be Registered
         */
        [TestMethod]
        public void TestRegisterListener()
        {
            if (factory.listeners.Count != 0)
            {
                factory = new ModelFactory();
            }

            factory.RegisterListener(new MFListener(null));
            factory.RegisterListener(new MFListener(null));
            Assert.IsTrue(
                        factory.listeners.Count == 2,
                        "Factory listeners should be two, was=" +
                                      factory.listeners.Count);
        }

        /**
         * Tests that given a simple input stream with one workbook and sheet
         * that those models are Processed and returned.
         */
        [TestMethod]
        public void TestRun()
        {
            Model temp = null;
            IEnumerator mi = null;

            if (factory.listeners.Count != 0)
            {
                factory = new ModelFactory();
            }

            factory.RegisterListener(new MFListener(models));
            factory.Run(in1);

            Assert.IsTrue(models.Count == 2, "Models size must be 2 was = " + models.Count);
            mi = models.GetEnumerator();
            mi.MoveNext();
            temp = (Model)mi.Current;

            Assert.IsTrue(
                temp is InternalWorkbook,
                "First model is Workbook was " + temp.GetType().Name);

            mi.MoveNext();
            temp = (Model)mi.Current;

            Assert.IsTrue(
                        temp is InternalSheet,
                        "Second model is Sheet was " + temp.GetType().Name);

        }

        /**
         * Sets up a Test file
         */
        private MemoryStream SetupRunFile(HSSFWorkbook book)
        {
            MemoryStream stream = new MemoryStream();
            NPOI.SS.UserModel.ISheet sheet = book.CreateSheet("Test");
            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell.SetCellValue(10.5);
            book.Write(stream);
            return stream;
        }

    }

    /**
     * listener for use in the Test
     */
    class MFListener : ModelFactoryListener
    {
        private ArrayList mlist;
        public MFListener(ArrayList mlist)
        {
            this.mlist = mlist;
        }

        public bool Process(Model model)
        {
            mlist.Add(model);
            return true;
        }

        public IEnumerator models()
        {
            return mlist.GetEnumerator();
        }

    }
}