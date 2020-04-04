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

using System.IO;
using System;

using System.Collections.Generic;
using ICSharpCode.SharpZipLib.Zip;
using NUnit.Framework;
namespace TestCases.OpenXml4Net.OPC
{

    /**
     * Compare the contents of 2 zip files.
     * 
     * @author CDubettier
     */
    public class ZipFileAssert
    {


        static int BUFFER_SIZE = 2048;

        protected static bool Equals(
                Dictionary<String, MemoryStream> file1,
                Dictionary<String, MemoryStream> file2)
        {
            Dictionary<String, MemoryStream>.KeyCollection listFile1 = file1.Keys;
            if (listFile1.Count == file2.Keys.Count)
            {
                for (Dictionary<String, MemoryStream>.KeyCollection.Enumerator iter = listFile1.GetEnumerator();
                    iter.MoveNext(); )
                {
                    String fileName = (String)iter.Current;
                    // extract the contents for both
                    MemoryStream Contain2;
                    if (file2.ContainsKey(fileName))
                        Contain2 = file2[fileName];
                    else
                        Contain2 = null;

                    MemoryStream Contain1;
                    if (file1.ContainsKey(fileName))
                        Contain1 = file1[fileName];
                    else
                        Contain1 = null;


                    if (Contain2 == null)
                    {
                        // file not found in archive 2
                        Assert.Fail(fileName + " not found in 2nd zip");
                        return false;
                    }
                    // no need to check for Contain1. The key come from it

                    if ((fileName.EndsWith(".xml")) || fileName.EndsWith(".rels"))
                    {
                        // we have a xml file
                        // TODO
                        // YK: the original OpenXML4J version attempted to compare xml using xmlunit (http://xmlunit.sourceforge.net),
                        // but POI does not depend on this library
                    }
                    else
                    {
                        // not xml, may be an image or other binary format
                        if (Contain2.Length != Contain1.Length)
                        {
                            // not the same size
                            Assert.Fail(fileName
                                    + " does not have the same size in both zip:"
                                    + Contain2.Length + "!=" + Contain1.Length);
                            return false;
                        }
                        byte[] array1 = Contain1.ToArray();
                        byte[] array2 = Contain2.ToArray();
                        for (int i = 0; i < array1.Length; i++)
                        {
                            if (array1[i] != array2[i])
                            {
                                Assert.Fail(fileName + " differ at index:" + i);
                                return false;
                            }
                        }
                    }
                }
            }
            else
            {
                // not the same number of files -> cannot be Equals
                Assert.Fail("not the same number of files in zip:"
                        + listFile1.Count + "!=" + file2.Keys.Count);
                return false;
            }
            return true;
        }

        protected static Dictionary<String, MemoryStream> decompress(
                FileInfo filename)
        {
            // store the zip content in memory
            // let s assume it is not Go ;-)
            Dictionary<String, MemoryStream> zipContent = new Dictionary<String, MemoryStream>();

            byte[] data = new byte[BUFFER_SIZE];
            /* Open file to decompress */
            FileStream file_decompress = filename.OpenRead();

            
            /* Open the file with the buffer */
            ZipInputStream zis = new ZipInputStream(file_decompress);

            /* Processing entries of the zip file */
            ZipEntry entree;
            int count;
            while ((entree = zis.GetNextEntry()) != null)
            {

                /* Create a array for the current entry */
                MemoryStream byteArray = new MemoryStream();
                zipContent.Add(entree.Name, byteArray);

                /* copy in memory */
                while ((count = zis.Read(data, 0, BUFFER_SIZE)) != 0)
                {
                    byteArray.Write(data, 0, count);
                }
                /* Flush the buffer */
                byteArray.Flush();
                //byteArray.Close();
            }

            zis.Close();

            return zipContent;
        }

        /**
         * Asserts that two files are Equal. Throws an <tt>AssertionFailedError</tt>
         * if they are not.
         * <p>
         * 
         */
        public static void AssertEqual(FileInfo expected, FileInfo actual)
        {
            Assert.IsNotNull(expected);
            Assert.IsNotNull(actual);

            Assert.IsTrue(File.Exists(expected.FullName), "File does not exist [" + expected.FullName
                    + "]");
            Assert.IsTrue(File.Exists(actual.FullName), "File does not exist [" + actual.FullName
                    + "]");

            //Assert.IsTrue("Expected file not Readable", expected.anRead());
            //Assert.IsTrue("Actual file not Readable", actual.canRead());

            try
            {
                Dictionary<String, MemoryStream> file1 = decompress(expected);
                Dictionary<String, MemoryStream> file2 = decompress(actual);
                Equals(file1, file2);
            }
            catch (IOException e)
            {
                throw new AssertionException(e.ToString());
            }
        }
    }
}




