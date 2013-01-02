
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
    using System.Collections;
    using NUnit.Framework;
    using NPOI.HPSF;
    using NPOI.POIFS.EventFileSystem;
    using NPOI.Util;



    /**
     * Static utility methods needed by the HPSF Test cases.
     *
     * @author Rainer Klute (klute@rainer-klute.de)
     * @since 2002-07-20
     * @version $Id: Util.java 489730 2006-12-22 19:18:16Z bayard $
     */
    public class Util
    {

        /**
         * Reads bytes from an input stream and Writes them to an
         * output stream until end of file is encountered.
         *
         * @param in the input stream to Read from
         * 
         * @param out the output stream to Write to
         * 
         * @exception IOException if an I/O exception occurs
         */
        public static void Copy(Stream in1, Stream out1)
        {
            int BUF_SIZE = 1000;
            byte[] b = new byte[BUF_SIZE];
            int Read;
            bool eof = false;
            while (!eof)
            {
                try
                {
                    Read = in1.Read(b, 0, BUF_SIZE);
                    if (Read > 0)
                        out1.Write(b, 0, Read);
                    else
                        eof = true;
                }
                catch 
                {
                    eof = true;
                }
            }
        }



        /**
         * Reads all files from a POI filesystem and returns them as an
         * array of {@link POIFile} instances. This method loads all files
         * into memory and thus does not cope well with large POI
         * filessystems.
         * 
         * @param poiFs The name of the POI filesystem as seen by the
         * operating system. (This is the "filename".)
         *
         * @return The POI files. The elements are ordered in the same way
         * as the files in the POI filesystem.
         * 
         * @exception FileNotFoundException if the file containing the POI 
         * filesystem does not exist
         * 
         * @exception IOException if an I/O exception occurs
         */
        public static POIFile[] ReadPOIFiles(Stream poiFs)
        {
            return ReadPOIFiles(poiFs, null);
        }

        /**
         * Reads a Set of files from a POI filesystem and returns them
         * as an array of {@link POIFile} instances. This method loads all
         * files into memory and thus does not cope well with large POI
         * filessystems.
         * 
         * @param poiFs The name of the POI filesystem as seen by the
         * operating system. (This is the "filename".)
         *
         * @param poiFiles The names of the POI files to be Read.
         *
         * @return The POI files. The elements are ordered in the same way
         * as the files in the POI filesystem.
         * 
         * @exception FileNotFoundException if the file containing the POI 
         * filesystem does not exist
         * 
         * @exception IOException if an I/O exception occurs
         */
        public static POIFile[] ReadPOIFiles(Stream poiFs,
                                             String[] poiFiles)
        {
            files = new ArrayList();
            POIFSReader reader1 = new POIFSReader();
            //reader1.StreamReaded += new POIFSReaderEventHandler(reader1_StreamReaded);
            POIFSReaderListener pfl = new POIFSReaderListener0();
            if (poiFiles == null)
                /* Register the listener for all POI files. */
                reader1.RegisterListener(pfl);
            else
                /* Register the listener for the specified POI files
                 * only. */
                for (int i = 0; i < poiFiles.Length; i++)
                    reader1.RegisterListener(pfl, poiFiles[i]);

            /* Read the POI filesystem. */
            reader1.Read(poiFs);
            POIFile[] result = new POIFile[files.Count];
            for (int i = 0; i < result.Length; i++)
                result[i] = (POIFile)files[i];
            return result;
        }
        private class POIFSReaderListener0 : POIFSReaderListener
        {
            #region POIFSReaderListener members

            public void ProcessPOIFSReaderEvent(POIFSReaderEvent evt)
            {
                try
                {
                    POIFile f = new POIFile();
                    f.SetName(evt.Name);
                    f.SetPath(evt.Path);
                    MemoryStream out1 =
                        new MemoryStream();
                    Util.Copy(evt.Stream, out1);
                    out1.Close();
                    f.SetBytes(out1.ToArray());
                    files.Add(f);
                }
                catch (IOException ex)
                {
                    throw new RuntimeException(ex.Message);
                }
            }

            #endregion
        }

        
        static IList files;
        /**
         * Read all files from a POI filesystem which are property Set streams
         * and returns them as an array of {@link org.apache.poi.hpsf.PropertySet}
         * instances.
         * 
         * @param poiFs The name of the POI filesystem as seen by the
         * operating system. (This is the "filename".)
         *
         * @return The property Sets. The elements are ordered in the same way
         * as the files in the POI filesystem.
         * 
         * @exception FileNotFoundException if the file containing the POI 
         * filesystem does not exist
         * 
         * @exception IOException if an I/O exception occurs
         */
        public static POIFile[] ReadPropertySets(FileStream poifs)
        {
            files = new ArrayList(7);
            POIFSReader reader2 = new POIFSReader();
            //reader2.StreamReaded += new POIFSReaderEventHandler(reader2_StreamReaded);
            POIFSReaderListener pfl = new POIFSReaderListener1();
            reader2.RegisterListener(pfl);
            /* Read the POI filesystem. */
            reader2.Read(poifs);
            POIFile[] result = new POIFile[files.Count];
            for (int i = 0; i < result.Length; i++)
                result[i] = (POIFile)files[i];
            return result;
        }
        private class POIFSReaderListener1:POIFSReaderListener
        {
            #region POIFSReaderListener members

            public void ProcessPOIFSReaderEvent(POIFSReaderEvent e)
            {
                try
                {
                    POIFile f = new POIFile();
                    f.SetName(e.Name);
                    f.SetPath(e.Path);
                    Stream in1 = e.Stream;
                    if (PropertySet.IsPropertySetStream(in1))
                    {
                        using (MemoryStream out1 = new MemoryStream())
                        {
                            Util.Copy(in1, out1);
                            //out1.Close();
                            f.SetBytes(out1.ToArray());
                            files.Add(f);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new RuntimeException(ex);
                }
            }

            #endregion
        }




        /**
         * Prints the system properties to System.out.
         */
        public static void PrintSystemProperties()
        {
            IDictionary p = Environment.GetEnvironmentVariables();
            ArrayList names = new ArrayList();
            for (IEnumerator i = p.Keys.GetEnumerator(); i.MoveNext(); )
                names.Add(i.Current);
            names.Sort();
            for (IEnumerator i = names.GetEnumerator(); i.MoveNext(); )
            {
                String name = (String)i.Current;
                String value = (String)p[name];
                Console.WriteLine(name + ": " + value);
            }
            Console.WriteLine("Current directory: " +
                               Environment.GetEnvironmentVariable("user.dir"));
        }

    }
}
