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
    /// EventArgs for POIFSReader
    /// author: Tony Qu
    /// </summary>
    public class POIFSReaderEventArgs:EventArgs
    {
        public POIFSReaderEventArgs(string name, POIFSDocumentPath path, POIFSDocument document)
        {
            this.name = name;
            this.path = path;
            this.document = document;
        }

        private POIFSDocumentPath path;
        private POIFSDocument document;
        private string name;

        public virtual POIFSDocumentPath Path
        {
            get { return path; }
        }
        public virtual POIFSDocument Document
        {
            get { return document; }
        }
        public virtual DocumentInputStream Stream
        {
            get { 
                return new DocumentInputStream(this.document); 
            }
        }
        public virtual string Name
        {
            get { return name; }
        }
    }

    public delegate void POIFSReaderEventHandler(object sender, POIFSReaderEventArgs e);
}
