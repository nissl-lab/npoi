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


namespace TestCases.POIFS.FileSystem
{
    /**
     * Class to Test POIFSDocumentPath functionality
     *
     * @author Marc Johnson
     */
    [TestFixture]
    public class TestPOIFSDocumentPath
    {

        /**
         * Constructor TestPOIFSDocumentPath
         *
         * @param name
         */

        public TestPOIFSDocumentPath()
        {

        }

        /**
         * Test default constructor
         */
        [Test]
        public void TestDefaultConstructor()
        {
            POIFSDocumentPath path = new POIFSDocumentPath();

            Assert.AreEqual(0, path.Length);
        }

        /**
         * Test full path constructor
         */
        [Test]
        public void TestFullPathConstructor()
        {
            string[] components = { "foo", "bar", "foobar", "fubar" };

            for (int j = 0; j < components.Length; j++)
            {
                string[] pms = new string[j];

                for (int k = 0; k < j; k++)
                {
                    pms[k] = components[k];
                }

                POIFSDocumentPath path = new POIFSDocumentPath(pms);

                Assert.AreEqual(j, path.Length);
                for (int k = 0; k < j; k++)
                {
                    Assert.AreEqual(components[k], path.GetComponent(k));
                }
                if (j == 0)
                    Assert.IsNull(path.Parent);
                else
                {
                    POIFSDocumentPath parent = path.Parent;

                    Assert.IsNotNull(parent);
                    Assert.AreEqual(j - 1, parent.Length);
                    for (int k = 0; k < j - 1; k++)
                        Assert.AreEqual(components[k], parent.GetComponent(k));
                }
            }

            // Test weird variants
            Assert.AreEqual(0, new POIFSDocumentPath(null).Length);
            try
            {
                new POIFSDocumentPath(new string[] { "fu", "" });
                Assert.Fail("Should have caught IllegalArgumentException");
            }
            catch (ArgumentException)
            {
            }
            try
            {
                new POIFSDocumentPath(new string[] { "fu", null });
                Assert.Fail("Should have caught IllegalArgumentException");
            }
            catch (ArgumentException)
            {
            }
        }

        /**
         * Test relative path constructor
         */
        [Test]
        public void TestRelativePathConstructor()
        {
            string[] initialComponents = { "a", "b", "c" };

            for (int n = 0; n < initialComponents.Length; n++)
            {
                String[] initialParams = new String[n];

                for (int k = 0; k < n; k++)
                    initialParams[k] = initialComponents[k];

                POIFSDocumentPath b = new POIFSDocumentPath(initialParams);
                string[] components = { "foo", "bar", "foobar", "fubar" };

                for (int j = 0; j < components.Length; j++)
                {
                    String[] params1 = new String[j];

                    for (int k = 0; k < j; k++)
                    {
                        params1[k] = components[k];
                    }
                    POIFSDocumentPath path = new POIFSDocumentPath(b, params1);

                    Assert.AreEqual(j + n, path.Length);
                    for (int k = 0; k < n; k++)
                    {
                        Assert.AreEqual(initialComponents[k],
                                     path.GetComponent(k));
                    }
                    for (int k = 0; k < j; k++)
                    {
                        Assert.AreEqual(components[k], path.GetComponent(k + n));
                    }
                    if ((j + n) == 0)
                    {
                        Assert.IsNull(path.Parent);
                    }
                    else
                    {
                        POIFSDocumentPath parent = path.Parent;

                        Assert.IsNotNull(parent);
                        Assert.AreEqual(j + n - 1, parent.Length);
                        for (int k = 0; k < (j + n - 1); k++)
                        {
                            Assert.AreEqual(path.GetComponent(k), parent.GetComponent(k));
                        }
                    }
                }

                Assert.AreEqual(n, new POIFSDocumentPath(b, null).Length);
                //this one is allowed.
                new POIFSDocumentPath(b, new string[] { "fu", "" });

                //this one is allowed too
                new POIFSDocumentPath(b, new string[] { "", "fu" });

                //this one is not allowed.
                try
                {
                    new POIFSDocumentPath(b, new string[] { "fu", null });
                    Assert.Fail("should have caught ArgumentException");
                }
                catch (ArgumentException)
                {
                }
                try
                {
                    new POIFSDocumentPath(b, new string[] { "fu", null });
                    Assert.Fail("should have caught ArgumentException");
                }
                catch (ArgumentException)
                {
                }
            }
        }

        /**
         * Test equality
         */
        [Test]
        public void TestEquality()
        {
            POIFSDocumentPath a1 = new POIFSDocumentPath();
            POIFSDocumentPath a2 = new POIFSDocumentPath(null);
            POIFSDocumentPath a3 = new POIFSDocumentPath(new String[0]);
            POIFSDocumentPath a4 = new POIFSDocumentPath(a1, null);
            POIFSDocumentPath a5 = new POIFSDocumentPath(a1, new string[0]);

            POIFSDocumentPath[] paths = { a1, a2, a3, a4, a5 };

            for (int j = 0; j < paths.Length; j++)
            {
                for (int k = 0; k < paths.Length; k++)
                    Assert.AreEqual(paths[j], paths[k], j + "<>" + k);
            }

            a2 = new POIFSDocumentPath(a1, new string[] { "foo" });
            a3 = new POIFSDocumentPath(a2, new string[] { "bar" });
            a4 = new POIFSDocumentPath(a3, new string[] { "fubar" });
            a5 = new POIFSDocumentPath(a4, new string[] { "foobar" });

            POIFSDocumentPath[] builtUpPaths = { a1, a2, a3, a4, a5 };

            POIFSDocumentPath[] fullPaths =
            {
                    new POIFSDocumentPath(),
                    new POIFSDocumentPath(new string[]{"foo"}),
                    new POIFSDocumentPath(new string[]{"foo", "bar"}),
                    new POIFSDocumentPath(new string[]{"foo", "bar", "fubar"}),
                    new POIFSDocumentPath(new string[]{"foo", "bar", "fubar", "foobar"})
            };

            for (int k = 0; k < builtUpPaths.Length; k++)
            {
                for (int j = 0; j < fullPaths.Length; j++)
                {
                    if (k == j)
                        Assert.AreEqual(fullPaths[j], builtUpPaths[k], j + "<>" + k);
                    else
                        Assert.IsTrue(!(fullPaths[j].Equals(builtUpPaths[k])), j + "<>" + k);
                }
            }
            POIFSDocumentPath[] badPaths =
            {
                    new POIFSDocumentPath(new string[]{"_foo"}),
                    new POIFSDocumentPath(new string[]{"foo", "_bar"}),
                    new POIFSDocumentPath(new string[]{"foo", "bar", "_fubar"}),
                    new POIFSDocumentPath(new string[]{"foo", "bar", "fubar", "_foobar"})

            };

            for (int k = 0; k < builtUpPaths.Length; k++)
            {
                for (int j = 0; j < badPaths.Length; j++)
                    Assert.IsTrue(!(fullPaths[k].Equals(badPaths[j])), j + "<>" + k);
            }
        }
    }
}
