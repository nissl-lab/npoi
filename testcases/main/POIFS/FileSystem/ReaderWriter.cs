
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
        
using System;
using System.Collections;
using System.IO;


using NPOI.POIFS.EventFileSystem;
using NPOI.POIFS.FileSystem;

/**
 * Test (Proof of concept) program that employs the
 * POIFSReaderListener and POIFSWriterListener interfaces
 *
 * @author Marc Johnson (mjohnson at apache dot org)
 */
namespace TestCases.POIFS.FileSystem
{
    public class ReaderWriter : POIFSReaderListener, POIFSWriterListener
    {
        private POIFSFileSystem filesystem;
        private DirectoryEntry root;

        // keys are DocumentDescriptors, values are byte[]s
        private Hashtable dataMap;

        public ReaderWriter(POIFSFileSystem fileSystem)
        {
            this.filesystem = fileSystem;
            root = this.filesystem.Root;
            dataMap = new Hashtable();
        }

        /* ********** START implementation of POIFSReaderListener ********** */

        /**
         * Process a POIFSReaderEvent that this listener had registered
         * for
         *
         * @param evt the POIFSReaderEvent
         */

        public void ProcessPOIFSReaderEvent(POIFSReaderEvent evt)
        {
            DocumentInputStream istream = evt.Stream;
            POIFSDocumentPath path = evt.Path;
            String name = evt.Name;

            try
            {
                int size = (int)(istream.Length - istream.Position);
                byte[] data = new byte[size];

                istream.Read(data);
                DocumentDescriptor descriptor = new DocumentDescriptor(path,
                                                    name);

                Console.WriteLine("Adding document: " + descriptor + " (" + size
                                   + " bytes)");
                dataMap[descriptor] = data;
                DirectoryEntry entry = root;

                for (int k = 0; k < path.Length; k++)
                {
                    String componentName = path.GetComponent(k);
                    Entry nextEntry = null;

                    try
                    {
                        nextEntry = entry.GetEntry(componentName);
                    }
                    catch (FileNotFoundException)
                    {
                        try
                        {
                            nextEntry = entry.CreateDirectory(componentName);
                        }
                        catch (IOException)
                        {
                            Console.WriteLine("Unable to Create directory");
                            //e.printStackTrace();
                            throw;
                        }
                    }
                    entry = (DirectoryEntry)nextEntry;
                }
                entry.CreateDocument(name, size, this);
            }
            catch (IOException)
            {
            }
        }

        /* **********  END  implementation of POIFSReaderListener ********** */
        /* ********** START implementation of POIFSWriterListener ********** */

        /**
         * Process a POIFSWriterEvent that this listener had registered
         * for
         *
         * @param evt the POIFSWriterEvent
         */

        public void ProcessPOIFSWriterEvent(POIFSWriterEvent evt)
        {
            try
            {
                DocumentDescriptor descriptor =
                    new DocumentDescriptor(evt.Path, evt.Name);

                Console.WriteLine("looking up document: " + descriptor + " ("
                                   + evt.Limit + " bytes)");
                evt.Stream.Write((byte[])dataMap[descriptor]);
            }
            catch (IOException)
            {
                Console.WriteLine("Unable to Write document");
                //e.printStackTrace();
                //System.exit(1);
                throw;
            }
        }

        /* **********  END  implementation of POIFSWriterListener ********** */
    }   // end public class ReaderWriter

}