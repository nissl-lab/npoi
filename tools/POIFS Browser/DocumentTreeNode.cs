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

/* ================================================================
 * POIFS Browser 
 * Author: NPOI Team
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * Huseyin Tufekcilerli     2008.11         1.0
 * Tony Qu                  2009.2.18       1.2 alpha
 * 
 * ==============================================================*/

using NPOI.POIFS.FileSystem;
using System.IO;
using NPOI.Util;

namespace NPOI.Tools.POIFSBrowser
{
    internal class DocumentTreeNode : AbstractTreeNode
    {
        public DocumentNode DocumentNode { get; private set; }

        public DocumentTreeNode(DocumentNode documentNode)
            : base(documentNode)
        {
            this.DocumentNode = documentNode;

            ChangeImage(documentNode.Name.EndsWith("SummaryInformation") ? "SummaryStream" 
                                                                         : "File");
        }

        public Stream GetDocumentStream()
        {
            var document = this.DocumentNode.Document;

            var dst = new byte[document.Size];

            if (document.Size > 0)
            {
                document.Read(dst, 0);
            }

            return new ByteArrayInputStream(dst);
        }
    }
}