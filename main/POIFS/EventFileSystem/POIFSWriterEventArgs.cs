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
using NPOI.POIFS.FileSystem;

namespace NPOI.POIFS.EventFileSystem
{


    /// <summary>
    /// EventArgs for POIFSWriter
    /// author: Tony Qu
    /// </summary>
    public class POIFSWriterEventArgs : EventArgs
    {

        private string documentName;
        private int limit;
        private POIFSDocumentPath path;
        private DocumentOutputStream stream;

        /// <summary>
        /// Initializes a new instance of the <see cref="POIFSWriterEvent"/> class.
        /// </summary>
        /// <param name="stream">the POIFSDocumentWriter, freshly opened</param>
        /// <param name="path">the path of the document</param>
        /// <param name="documentName">the name of the document</param>
        /// <param name="limit">the limit, in bytes, that can be written to the stream</param>
        public POIFSWriterEventArgs(DocumentOutputStream stream, POIFSDocumentPath path, string documentName, int limit)
        {
            this.stream = stream;
            this.path = path;
            this.documentName = documentName;
            this.limit = limit;
        }
        /// <summary>
        /// Gets the limit on writing, in bytes
        /// </summary>
        /// <value>The limit.</value>
        public virtual int Limit
        {
            get
            {
                return this.limit;
            }
        }

        /// <summary>
        /// Gets the document's name
        /// </summary>
        /// <value>The name.</value>
        public virtual string Name
        {
            get
            {
                return this.documentName;
            }
        }

        /// <summary>
        /// Gets the document's path
        /// </summary>
        /// <value>The path.</value>
        public virtual POIFSDocumentPath Path
        {
            get
            {
                return this.path;
            }
        }

        /// <summary>
        /// the POIFSDocumentWriter, freshly opened
        /// </summary>
        /// <value>The stream.</value>
        public virtual DocumentOutputStream Stream
        {
            get
            {
                return this.stream;
            }
        }
    }

    public delegate void POIFSWriterEventHandler(object sender, POIFSWriterEventArgs e);
}
