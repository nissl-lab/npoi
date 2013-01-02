
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
    using NPOI.POIFS.FileSystem;


    /**
     * A POI file just for Testing.
     *
     * @author Rainer Klute (klute@rainer-klute.de)
     * @since 2002-07-20
     * @version $Id: POIFile.java 489730 2006-12-22 19:18:16Z bayard $
     */
    public class POIFile
    {

        private String name;
        private POIFSDocumentPath path;
        private byte[] bytes;


        /**
         * Sets the POI file's name.
         *
         * @param name The POI file's name.
         */
        public void SetName(String name)
        {
            this.name = name;
        }

        /**
         * Returns the POI file's name.
         *
         * @return The POI file's name.
         */
        public String GetName()
        {
            return name;
        }

        /**
         * Sets the POI file's path.
         *
         * @param path The POI file's path.
         */
        public void SetPath(POIFSDocumentPath path)
        {
            this.path = path;
        }

        /**
         * Returns the POI file's path.
         *
         * @return The POI file's path.
         */
        public POIFSDocumentPath GetPath()
        {
            return path;
        }

        /**
         * Sets the POI file's content bytes.
         *
         * @param bytes The POI file's content bytes.
         */
        public void SetBytes(byte[] bytes)
        {
            this.bytes = bytes;
        }

        /**
         * Returns the POI file's content bytes.
         *
         * @return The POI file's content bytes.
         */
        public byte[] GetBytes()
        {
            return bytes;
        }

    }
}