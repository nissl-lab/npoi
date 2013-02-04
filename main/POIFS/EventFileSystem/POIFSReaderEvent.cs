
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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
 * Class POIFSReaderEvent
 *
 * @author Marc Johnson (mjohnson at apache dot org)
 * @version %I%, %G%
 */
    public class POIFSReaderEvent
    {
        private DocumentInputStream stream;
        private POIFSDocumentPath path;
        private String documentName;

        /**
         * package scoped constructor
         *
         * @param stream the DocumentInputStream, freshly opened
         * @param path the path of the document
         * @param documentName the name of the document
         */

        public POIFSReaderEvent(DocumentInputStream stream,
                         POIFSDocumentPath path, String documentName)
        {
            this.stream = stream;
            this.path = path;
            this.documentName = documentName;
        }

        /**
         * @return the DocumentInputStream, freshly opened
         */

        public DocumentInputStream Stream
        {
            get { return stream; }
        }

        /**
         * @return the document's path
         */

        public POIFSDocumentPath Path
        {
            get { return path; }
        }

        /**
         * @return the document's name
         */

        public String Name
        {
            get { return documentName; }
        }
    }   // end public class POIFSReaderEvent

}