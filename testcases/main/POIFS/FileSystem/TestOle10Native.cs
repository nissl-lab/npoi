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

namespace TestCases.POIFS.FileSystem
{
    using System;
    using NUnit.Framework;
    using TestCases;
    using System.IO;
    using System.Collections.Generic;
    using NPOI.Util;
    using NPOI.POIFS.FileSystem;



    [TestFixture]
    public class TestOle10Native
    {
        private static POIDataSamples dataSamples = POIDataSamples.GetPOIFSInstance();


        [Test]
        public void TestOleNative()
        {
            POIFSFileSystem fs = new POIFSFileSystem(dataSamples.OpenResourceAsStream("oleObject1.bin"));

            Ole10Native ole = Ole10Native.CreateFromEmbeddedOleObject(fs);

            Assert.AreEqual("File1.svg", ole.Label);
            Assert.AreEqual("D:\\Documents and Settings\\rsc\\My Documents\\file1.svg", ole.Command);
        }

        [Test]
        public void TestFiles()
        {
            FileStream[] files = {
            // bug 51891
            POIDataSamples.GetPOIFSInstance().GetFile("multimedia.doc"),
            // tika bug 1072
            POIDataSamples.GetPOIFSInstance().GetFile("20-Force-on-a-current-S00.doc"),
            // other files Containing ole10native records ...
            POIDataSamples.GetDocumentInstance().GetFile("Bug53380_3.doc"),
            POIDataSamples.GetDocumentInstance().GetFile("Bug47731.doc")
        };

            foreach (FileStream f in files)
            {
                NPOIFSFileSystem fs = new NPOIFSFileSystem(f,null, true, true);
                List<Entry> entries = new List<Entry>();
                FindOle10(entries, fs.Root, "/", "");

                foreach (Entry e in entries)
                {
                    MemoryStream bosExp = new MemoryStream();
                    Stream is1 = ((DirectoryNode)e.Parent).CreateDocumentInputStream(e);
                    IOUtils.Copy(is1, bosExp);
                    is1.Close();

                    Ole10Native ole = Ole10Native.CreateFromEmbeddedOleObject((DirectoryNode)e.Parent);

                    MemoryStream bosAct = new MemoryStream();
                    ole.WriteOut(bosAct);

                    //assertThat(bosExp.ToByteArray(), EqualTo(bosAct.ToByteArray()));
                    Assert.IsTrue(Arrays.Equals(bosExp.ToArray(), bosAct.ToArray()));
                }

                fs.Close();
            }
        }

        /*
        void searchOle10Files()  {
            File dir = new File("test-data/document");
            foreach (File file in dir.ListFiles(new FileFilter(){
                public bool accept(File pathname) {
                    return pathname.Name.EndsWith("doc");
                }
            })) {
                NPOIFSFileSystem fs = new NPOIFSFileSystem(file, true);
                FindOle10(null, fs.Root, "/", file.Name);
                fs.Close();
            }
        }*/

        void FindOle10(List<Entry> entries, DirectoryNode dn, String path, String filename)
        {
            IEnumerator<Entry> iter = dn.Entries;
            while (iter.MoveNext())
            {
                Entry e = iter.Current;
                if (Ole10Native.OLE10_NATIVE.Equals(e.Name))
                {
                    if (entries != null) entries.Add(e);
                    // System.out.Println(filename+" : "+path);
                }
                else if (e.IsDirectoryEntry)
                {
                    FindOle10(entries, (DirectoryNode)e, path + e.Name + "/", filename);
                }
            }
        }
    }
}