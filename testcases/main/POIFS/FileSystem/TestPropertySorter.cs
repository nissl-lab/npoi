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

namespace TestCases.POIFS.FileSystem
{

    using System;
    using System.Collections;
    using System.IO;

    using NUnit.Framework;

    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using NPOI.POIFS.Storage;
    using NPOI.POIFS.Properties;

    using TestCases.HSSF;
    using System.Collections.Generic;

    /**
* Verify the order of entries <c>DirectoryProperty</c> .
* 
* In particular it is important to serialize ROOT._VBA_PROJECT_CUR.VBA node.
* See bug 39234 in bugzilla. Thanks to Bill Seddon for providing the solution.
* 
*
* @author Yegor Kozlov
*/
    [TestFixture]
    public class TestPropertySorter
    {
        private static IComparer<Property> OldCaseSensitivePropertyComparator = new PropertyComparer();



        //the correct order of entries in the Test file
        private static String[] _entries = {
            "dir", "JML", "UTIL", "Loader", "Sheet1", "Sheet2", "Sheet3",
            "__SRP_0", "__SRP_1", "__SRP_2", "__SRP_3", "__SRP_4", "__SRP_5",
            "ThisWorkbook","_VBA_PROJECT"
    };

        private static POIFSFileSystem OpenSampleFS()
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("39234.xls");
            try
            {
                return new POIFSFileSystem(is1);
            }
            catch (IOException)
            {
                throw;
            }
        }

        /**
         * Test sorting of properties in <c>DirectoryProperty</c>
         */
        [Test]
        public void TestSortProperties()
        {
            POIFSFileSystem fs = OpenSampleFS();
            Property[] props = GetVBAProperties(fs);

            Assert.AreEqual(_entries.Length, props.Length);

            // (1). See that there is a problem with the old case-sensitive property comparator
            Array.Sort(props, OldCaseSensitivePropertyComparator);
            //try
            //{
            //    for (int i = 0; i < props.Length; i++)
            //    {
            //        Assert.AreEqual(_entries[i], props[i].Name);
            //    }
            //    Assert.Fail("expected old case-sensitive property comparator to return properties in wrong order");
            //}
            //catch (AssertionException e)
            //{
            //    // expected during successful Test
            //    Assert.IsNotNull(e.Message);
            //}

            // (2) Verify that the fixed property comparator works right
            Array.Sort(props, new DirectoryProperty.PropertyComparator());
            for (int i = 0; i < props.Length; i++)
            {
                Assert.AreEqual(_entries[i], props[i].Name);
            }
        }

        /**
         * Serialize file system and Verify that the order of properties is the same as in the original file.
         */
        [Test]
        public void TestSerialization()
        {
            POIFSFileSystem fs = OpenSampleFS();

            MemoryStream out1 = new MemoryStream();
            fs.WriteFileSystem(out1);
            out1.Close();
            Stream is1 = new MemoryStream(out1.ToArray());
            fs = new POIFSFileSystem(is1);
            is1.Close();
            Property[] props = GetVBAProperties(fs);
            Array.Sort(props, new DirectoryProperty.PropertyComparator());

            Assert.AreEqual(_entries.Length, props.Length);
            for (int i = 0; i < props.Length; i++)
            {
                Assert.AreEqual(_entries[i], props[i].Name);
            }
        }

        /**
         * @return array of properties Read from ROOT._VBA_PROJECT_CUR.VBA node
         */
        protected Property[] GetVBAProperties(POIFSFileSystem fs)
        {
            String _VBA_PROJECT_CUR = "_VBA_PROJECT_CUR";
            String VBA = "VBA";

            DirectoryEntry root = fs.Root;
            DirectoryEntry vba_project = (DirectoryEntry)root.GetEntry(_VBA_PROJECT_CUR);

            DirectoryNode vba = (DirectoryNode)vba_project.GetEntry(VBA);
            DirectoryProperty p = (DirectoryProperty)vba.Property;

            ArrayList lst = new ArrayList();
            for (IEnumerator it = p.Children; it.MoveNext(); )
            {
                Property ch = (Property)it.Current;
                lst.Add(ch);
            }
            return (Property[])lst.ToArray(typeof(Property));
        }

        private class PropertyComparer : IComparer<Property>
        {
            public int Compare(Property o1, Property o2)
            {
                String name1 = o1.Name;
                String name2 = o2.Name;
                int result = name1.Length - name2.Length;

                if (result == 0)
                {
                    //result = name1.CompareTo(name2);
                    result = compareTo(name1, name2);
                }
                return result;
            }

            private int compareTo(string value, string anotherString)
            {
                int len1 = value.Length;
                int len2 = anotherString.Length;
                int lim = Math.Min(len1, len2);
                char[] v1 = value.ToCharArray();
                char[] v2 = anotherString.ToCharArray();

                int k = 0;
                while (k < lim)
                {
                    char c1 = v1[k];
                    char c2 = v2[k];
                    if (c1 != c2)
                    {
                        return c1 - c2;
                    }
                    k++;
                }
                return len1 - len2;
            }
        }
    }

}