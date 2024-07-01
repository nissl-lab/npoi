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

using NUnit.Framework;

using NPOI.POIFS.FileSystem;
using NPOI.Util;

using NPOI.POIFS.Storage;
using NPOI.POIFS.Properties;
using NPOI.POIFS.EventFileSystem;


namespace TestCases.POIFS.FileSystem
{
    [TestFixture]
    public class TestEmptyDocument
    {
        private static POILogger _logger = POILogFactory.GetLogger(typeof(TestEmptyDocument));

        private class AnonymousClass : POIFSWriterListener
        {
            public void ProcessPOIFSWriterEvent(POIFSWriterEvent ev)
            {
                TestEmptyDocument._logger.Log(POILogger.WARN, "Written");
                Console.WriteLine("Written");
            }
        }

        private class AnonymousClass1 :POIFSWriterListener
        {
            public void ProcessPOIFSWriterEvent(POIFSWriterEvent ev)
            {
                try
                {
                    ev.Stream.Write(0);
                    Console.WriteLine("Send an event");
                }
                catch (IOException exception)
                {
                    throw new Exception("Exception on write: " + exception.Message);
                }
            }
        }

        private class EmptyClass : POIFSWriterListener
        {
            public void ProcessPOIFSWriterEvent(POIFSWriterEvent ev)
            {
                try
                {
                    ;
                }
                catch (IOException exception)
                {
                    throw new Exception("Exception on write: " + exception.Message);
                }
            }
        }
       

        [Test]
        public void TestSingleEmptyDocument()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dir = fs.Root;
            dir.CreateDocument("Foo", new MemoryStream(new byte[] { }));

            MemoryStream output = new MemoryStream();
            fs.WriteFileSystem(output);
            
            new POIFSFileSystem(new ByteArrayInputStream(output.ToArray())).Close();
            fs.Close();
        }

        [Test]
        public void TestSingleEmptyDocumentEvent()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dir = fs.Root;
            dir.CreateDocument("Foo", 0, new AnonymousClass());

            MemoryStream output = new MemoryStream();
            fs.WriteFileSystem(output);
            new POIFSFileSystem(new ByteArrayInputStream(output.ToArray())).Close();
            fs.Close();
        }
        [Test]
        public void TestEmptyDocumentWithFriend()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dir = fs.Root;
            dir.CreateDocument("Bar", new MemoryStream(new byte[]{0}));
            dir.CreateDocument("Foo", new MemoryStream(new byte[]{}));

            MemoryStream output = new MemoryStream();
            fs.WriteFileSystem(output);
            new POIFSFileSystem(new ByteArrayInputStream(output.ToArray())).Close();
            fs.Close();
        }

        [Test]
        public void TestEmptyDocumentEventWithFriend()
        {
            POIFSFileSystem fs = new POIFSFileSystem();
            DirectoryEntry dir = fs.Root;
            dir.CreateDocument("Bar", 1, new AnonymousClass1());
            dir.CreateDocument("Foo", 0, new EmptyClass());

            MemoryStream output = new MemoryStream();
            fs.WriteFileSystem(output);
            new POIFSFileSystem(new ByteArrayInputStream(output.ToArray())).Close();
            fs.Close();

        }
        [Test]
        public void TestEmptyDocumentBug11744()
        {
            byte[] TestData = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };

            POIFSFileSystem fs = new POIFSFileSystem();
            fs.CreateDocument(new MemoryStream(new byte[0]), "Empty");
            fs.CreateDocument(new MemoryStream(TestData), "NotEmpty");
            MemoryStream output = new MemoryStream();
            fs.WriteFileSystem(output);
            fs.Close();

            // This line caused the error.
            fs = new POIFSFileSystem(new MemoryStream(output.ToArray()));

            DocumentEntry entry = (DocumentEntry)fs.Root.GetEntry("Empty");
            Assert.AreEqual(0, entry.Size, "Expected zero size");
            byte[] actualReadbackData;
            actualReadbackData = NPOI.Util.IOUtils.ToByteArray(new DocumentInputStream(entry));
            Assert.AreEqual(0, actualReadbackData.Length, "Expected zero read from stream");

            entry = (DocumentEntry)fs.Root.GetEntry("NotEmpty");
            actualReadbackData = NPOI.Util.IOUtils.ToByteArray(new DocumentInputStream(entry));
            Assert.AreEqual(TestData.Length, entry.Size, "Expected size was wrong");
            Assert.IsTrue(
                    Arrays.Equals(TestData,actualReadbackData), "Expected different data Read from stream");

            fs.Close();
        }
    }
}