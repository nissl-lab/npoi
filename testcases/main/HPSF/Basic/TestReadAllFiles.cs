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



namespace TestCases.HPSF.Basic
{
    using System;
    using System.IO;
    using System.Text;
    using System.Collections;
    using NUnit.Framework;
    using NPOI.HPSF;
    using NPOI.Util;
    using System.Collections.Generic;


    /**
     * Tests some HPSF functionality by Reading all property Sets from all files
     * in the "data" directory. If you want to ensure HPSF can deal with a certain
     * OLE2 file, just Add it to the "data" directory and run this Test case.
     * 
     * @author Rainer Klute (klute@rainer-klute.de)
     * @since 2008-02-08
     * @version $Id: TestBasic.java 489730 2006-12-22 19:18:16Z bayard $
     */
    [TestFixture]
    public class TestReadAllFiles
    {
        private static String[] excludes = new String[] {
        //"TestZeroLengthCodePage.mpp",
    };

        /**
         * Test case constructor.
         * 
         * @param name The Test case's name.
         */
        public TestReadAllFiles()
        {

        }



        /**
         * This Test methods Reads all property Set streams from all POI
         * filesystems in the "data" directory.
         */
        [Test]
        public void TestReadAllFiles1()
        {
            //string dataDir = @"..\..\..\TestCases\HPSF\data\";
            POIDataSamples _samples = POIDataSamples.GetHPSFInstance();
            string[] files = _samples.GetFiles();
            try
            {
                for (int i = 0; i < files.Length; i++)
                {
                    if (files[i].EndsWith("1") || !checkExclude(files[i]))
                        continue;

                    Console.WriteLine("Reading file \"" + files[i] + "\"");

                    using (FileStream f = new FileStream(files[i], FileMode.Open, FileAccess.Read))
                    {
                        /* Read the POI filesystem's property Set streams: */
                        POIFile[] psf1 = Util.ReadPropertySets(f);

                        for (int j = 0; j < psf1.Length; j++)
                        {
                            Stream in1 =
                                new ByteArrayInputStream(psf1[j].GetBytes());
                            try
                            {
                            PropertySet a = PropertySetFactory.Create(in1);
                            }
                            catch(Exception e)
                            {
                                throw new IOException("While handling file: " + files[i] + " at " + j, e);
                            }
                        }
                        f.Close();
                    }
                }
            }
            catch (Exception t)
            {
                String s = t.ToString();
                Assert.Fail(s);
            }
        }

        /**
         * Returns true if the file should be checked, false if it should be excluded.
         *
         * @param f
         * @return
         */
        public static bool checkExclude(string f)
        {
            foreach (String exclude in excludes)
            {
                if (f.EndsWith(exclude))
                {
                    return false;
                }
            }

            return true;
        }

    }
}