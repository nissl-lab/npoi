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
using System.Text;

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// Class DocumentDescriptor
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class DocumentDescriptor
    {
        private POIFSDocumentPath path;
        private String            name;
        private int               hashcode = 0;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentDescriptor"/> class.
        /// </summary>
        /// <param name="path">the Document path</param>
        /// <param name="name">the Document name</param>
        public DocumentDescriptor(POIFSDocumentPath path, String name)
        {
            if (path == null)
            {
                throw new NullReferenceException("path must not be null");
            }
            if (name == null)
            {
                throw new NullReferenceException("name must not be null");
            }
            if (name.Length== 0)
            {
                throw new ArgumentException("name cannot be empty");
            }
            this.path = path;
            this.name = name;
        }

        /// <summary>
        /// Gets the path.
        /// </summary>
        /// <value>The path.</value>
        public string Path
        {
            get { return this.path.ToString(); }
        }
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>The name.</value>
        public string Name
        {
            get { return this.name; }
        }

        /// <summary>
        /// equality. Two DocumentDescriptor instances are equal if they
        /// have equal paths and names
        /// </summary>
        /// <param name="o">the object we're checking equality for</param>
        /// <returns>true if the object is equal to this object</returns>
        public override bool Equals(Object o)
        {
            bool rval = false;

            if ((o != null) && (o.GetType()== this.GetType()))
            {
                if (this == o)
                {
                    rval = true;
                }
                else
                {
                    DocumentDescriptor descriptor = ( DocumentDescriptor ) o;

                    rval = this.path.Equals(descriptor.path)
                           && this.name.Equals(descriptor.name);
                }
            }
            return rval;
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// hashcode
        /// </returns>
        public override int GetHashCode()
        {
            if (hashcode == 0)
            {
                hashcode = path.GetHashCode() ^ name.GetHashCode();
            }
            return hashcode;
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override String ToString()
        {
            StringBuilder buffer = new StringBuilder(40 * (path.Length + 1));

            for (int j = 0; j < path.Length; j++)
            {
                buffer.Append(path.GetComponent(j)).Append("/");
            }
            buffer.Append(name);
            return buffer.ToString();
        }
    }
}
