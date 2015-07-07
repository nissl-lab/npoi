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

namespace NPOI.POIFS.Properties
{
    /// <summary>
    /// Trivial extension of Property for POIFSDocuments
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class DocumentProperty:Property
    {
       // the POIFSDocument this property is associated with
        private POIFSDocument _document;


        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentProperty"/> class.
        /// </summary>
        /// <param name="name">POIFSDocument name</param>
        /// <param name="size">POIFSDocument size</param>
        public DocumentProperty(String name, int size)
            : base()
        {
            _document = null;
            
            this.Name=name;
            this.Size=size;
            this.NodeColor=_NODE_BLACK;   // simplification
            this.PropertyType=PropertyConstants.DOCUMENT_TYPE;
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentProperty"/> class.
        /// </summary>
        /// <param name="index">index number</param>
        /// <param name="array">byte data</param>
        /// <param name="offset">offset into byte data</param> 
        public DocumentProperty(int index, byte [] array, int offset):base(index, array, offset)
        {
            
            _document = null;
        }

        /// <summary>
        /// Gets or sets the document.
        /// </summary>
        /// <value>the associated POIFSDocument</value>
        public POIFSDocument Document
        {
            set { _document = value; }
            get{ return _document;}
        }

        /// <summary>
        /// Determines whether this instance is directory.
        /// </summary>
        /// <returns>
        /// 	<c>true</c> if this instance is directory; otherwise, <c>false</c>.
        /// </returns>
        public override bool IsDirectory
        {
            get { return false; }
        }

        /// <summary>
        /// Perform whatever activities need to be performed prior to
        /// writing
        /// </summary>
        public override void PreWrite()
        {
            // do nothing
        }

        /**
         * Update the size of the property's data
         */
        public void UpdateSize(int size)
        {
            Size = (size);
        }
    }
}
