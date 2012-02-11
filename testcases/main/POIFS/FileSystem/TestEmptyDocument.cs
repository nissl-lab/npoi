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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.Collections;
using System.IO;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NPOI.POIFS.FileSystem;
using NPOI.Util;

using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;
using NPOI.POIFS.EventFileSystem;


namespace TestCases.POIFS.FileSystem
{
    [TestClass]
    public class TestEmptyDocument
    {

        public static void main(String[] args)
        {
            TestEmptyDocument driver = new TestEmptyDocument();

            Console.WriteLine();
            Console.WriteLine("As only file...");
            Console.WriteLine();

            Console.Write("Trying using CreateDocument(String,InputStream): ");
            try
            {
                driver.TestSingleEmptyDocument();
                Console.WriteLine("Worked!");
            }
            catch (IOException exception)
            {
                Console.WriteLine("Assert.Failed! ");
                Console.WriteLine(exception.ToString());
            }
            Console.WriteLine();

            Console.Write
              ("Trying using CreateDocument(String,int,POIFSWriterListener): ");
            try
            {
                driver.TestSingleEmptyDocumentEvent();
                Console.WriteLine("Worked!");
            }
            catch (IOException exception)
            {
                Console.WriteLine("Assert.Failed!");
                Console.WriteLine(exception.ToString());
            }
            Console.WriteLine();

            Console.WriteLine();
            Console.WriteLine("After another file...");
            Console.WriteLine();

            Console.Write("Trying using CreateDocument(String,InputStream): ");
            try
            {
                driver.TestEmptyDocumentWithFriend();
                Console.WriteLine("Worked!");
            }
            catch (IOException exception)
            {
                Console.WriteLine("Assert.Failed! ");
                Console.WriteLine(exception.ToString());
            }
            Console.WriteLine();

            Console.Write
              ("Trying using CreateDocument(String,int,POIFSWriterListener): ");
            try
            {
                driver.TestEmptyDocumentWithFriend();
                Console.WriteLine("Worked!");
            }
            catch (IOException exception)
            {
                Console.WriteLine("Assert.Failed!");
                Console.WriteLine(exception.ToString());
            }
            Console.WriteLine();
        }
        [TestMethod]
        public void TestSingleEmptyDocument()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dir = fs.Root;
            dir.CreateDocument("Foo", new MemoryStream(new byte[] { }));

            MemoryStream out1 = new MemoryStream();
            fs.WriteFileSystem(out1);
            new POIFSFileSystem(new MemoryStream(out1.ToArray()));
        }

        void OnWrite1(object sender,POIFSWriterEventArgs evt)
        {
            Console.WriteLine("written");
        }
        [TestMethod]
        public void TestSingleEmptyDocumentEvent()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dir = fs.Root;
            dir.CreateDocument("Foo", 0,new POIFSWriterEventHandler(OnWrite1));

            MemoryStream out1 = new MemoryStream();
            fs.WriteFileSystem(out1);
            new POIFSFileSystem(new MemoryStream(out1.ToArray()));
        }
        [TestMethod]
        public void TestEmptyDocumentWithFriend()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dir = fs.Root;
            dir.CreateDocument("Bar", new MemoryStream(new byte[] { 0 }));
            dir.CreateDocument("Foo", new MemoryStream(new byte[] { }));

            MemoryStream out1 = new MemoryStream();
            fs.WriteFileSystem(out1);
            new POIFSFileSystem(new MemoryStream(out1.ToArray()));
        }

            void OnWrite2(object sender,POIFSWriterEventArgs evt)
            {
                try
                {
                    evt.Stream.WriteByte(0);
                }
                catch (IOException exception)
                {
                    throw new InvalidOperationException("exception on Write: " + exception);
                }
            }


            void OnWrite3(object sender, POIFSWriterEventArgs evt)
            {

            }

        [TestMethod]
        public void TestEmptyDocumentEventWithFriend()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dir = fs.Root;

            dir.CreateDocument("Bar", 1,new POIFSWriterEventHandler(OnWrite2));
            dir.CreateDocument("Foo", 0, new POIFSWriterEventHandler(OnWrite3));

            MemoryStream out1 = new MemoryStream();
            fs.WriteFileSystem(out1);
            new POIFSFileSystem(new MemoryStream(out1.ToArray()));
        }
        [TestMethod]
        public void TestEmptyDocumentBug11744()
        {
            byte[] TestData = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };

            POIFSFileSystem fs = new POIFSFileSystem();
            fs.CreateDocument(new MemoryStream(new byte[0]), "Empty");
            fs.CreateDocument(new MemoryStream(TestData), "NotEmpty");
            MemoryStream out1 = new MemoryStream();
            fs.WriteFileSystem(out1);
            out1.ToArray();

            // This line caused the error.
            fs = new POIFSFileSystem(new MemoryStream(out1.ToArray()));

            DocumentEntry entry = (DocumentEntry)fs.Root.GetEntry("Empty");
            Assert.AreEqual(0, entry.Size, "Expected zero size");
            byte[] actualReadbackData = IOUtils.ToByteArray(new POIFSDocumentReader(entry));
            Assert.AreEqual(0,
                         actualReadbackData.Length, "Expected zero Read from stream");

            entry = (DocumentEntry)fs.Root.GetEntry("NotEmpty");
            actualReadbackData = IOUtils.ToByteArray(new POIFSDocumentReader(entry));
            Assert.AreEqual(TestData.Length, entry.Size, "Expected size was wrong");
            Assert.IsTrue(
                    Arrays.Equals(TestData,actualReadbackData), "Expected different data Read from stream");
        }
    }
}