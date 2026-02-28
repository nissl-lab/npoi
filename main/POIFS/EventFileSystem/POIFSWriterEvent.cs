
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
using NPOI.POIFS.FileSystem;
namespace NPOI.POIFS.EventFileSystem
{

    /**
     * Class POIFSWriterEvent
     *
     * @author Marc Johnson (mjohnson at apache dot org)
     * @version %I%, %G%
     */

    public class POIFSWriterEvent
    {
        private DocumentOutputStream stream;
        private POIFSDocumentPath path;
        private String documentName;
        private int limit;

        /**
         * namespace scoped constructor
         *
         * @param stream the DocumentOutputStream, freshly opened
         * @param path the path of the document
         * @param documentName the name of the document
         * @param limit the limit, in bytes, that can be written to the
         *              stream
         */

        public POIFSWriterEvent(DocumentOutputStream stream,
                         POIFSDocumentPath path, String documentName,
                         int limit)
        {
            this.stream = stream;
            this.path = path;
            this.documentName = documentName;
            this.limit = limit;
        }

        /**
         * @return the DocumentOutputStream, freshly opened
         */

        public DocumentOutputStream Stream
        {
            get
            {
                return stream;
            }
        }

        /**
         * @return the document's path
         */

        public POIFSDocumentPath Path
        {
            get
            {
                return path;
            }
        }

        /**
         * @return the document's name
         */

        public String Name
        {
            get
            {
                return documentName;
            }
        }

        /**
         * @return the limit on writing, in bytes
         */

        public int Limit
        {
            get
            {
                return limit;
            }
        }
    }   // end public class POIFSWriterEvent



}




