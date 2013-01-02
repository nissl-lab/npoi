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
using System.IO;

namespace NPOI.POIFS.FileSystem
{
    /// <summary>
    /// Class POIFSDocumentPath
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class POIFSDocumentPath
    {
        private string[] components;
        private int hashcode=0;

        /// <summary>
        /// simple constructor for the path of a document that is in the
        /// root of the POIFSFileSystem. The constructor that takes an
        /// array of Strings can also be used to create such a
        /// POIFSDocumentPath by passing it a null or empty String array
        /// </summary>
        public POIFSDocumentPath()
        {
            this.components = new string[0];
        }
        /// <summary>
        /// constructor for the path of a document that is not in the root
        /// of the POIFSFileSystem
        /// </summary>
        /// <param name="components">the Strings making up the path to a document.
        /// The Strings must be ordered as they appear in
        /// the directory hierarchy of the the document
        /// -- the first string must be the name of a
        /// directory in the root of the POIFSFileSystem,
        /// and every Nth (for N &gt; 1) string thereafter
        /// must be the name of a directory in the
        /// directory identified by the (N-1)th string.
        /// If the components parameter is null or has
        /// zero length, the POIFSDocumentPath is
        /// appropriate for a document that is in the
        /// root of a POIFSFileSystem</param>
        public POIFSDocumentPath(string[] components)
        {
            if (components == null)
            {
                this.components = new string[0];
            }
            else
            {
                this.components = new string[components.Length];
                for (int i = 0; i < components.Length; i++)
                {
                    if ((components[i] == null) 
                        || (components[i].Length == 0))
                    {
                        throw new ArgumentException("components cannot contain null or empty strings");
                    }
                    this.components[i] = components[i];
                }
            }
        }
        /// <summary>
        /// constructor that adds additional subdirectories to an existing
        /// path
        /// </summary>
        /// <param name="path">the existing path</param>
        /// <param name="components">the additional subdirectory names to be added</param>
        public POIFSDocumentPath(POIFSDocumentPath path, string[] components)
        {
            if (components == null)
            {
                this.components = new string[path.components.Length];
            }
            else
            {
                this.components = new string[path.components.Length + components.Length];
            }
            for (int i = 0; i < path.components.Length; i++)
            {
                this.components[i] = path.components[i];
            }
            if (components != null)
            {
                for (int j = 0; j < components.Length; j++)
                {
                    if (components[j] == null)
                    {
                        throw new ArgumentException("components cannot contain null");
                    }
                    if (components[j].Length == 0)
                    {
                       // throw new ArgumentException("components cannot contain null or empty strings");
                    }
                    this.components[j + path.components.Length] = components[j];
                }
            }
        }
        /// <summary>
        /// equality. Two POIFSDocumentPath instances are equal if they
        /// have the same number of component Strings, and if each
        /// component String is equal to its coresponding component String
        /// </summary>
        /// <param name="o">the object we're checking equality for</param>
        /// <returns>true if the object is equal to this object</returns>
        public override bool Equals(object o)
        {
            bool flag = false;
            if ((o != null) && (o.GetType() == this.GetType()))
            {
                if (this == o)
                {
                    flag = true;
                }
                else
                {
                    POIFSDocumentPath path = (POIFSDocumentPath)o;
                    if (path.components.Length == this.components.Length)
                    {
                        flag = true;
                        for (int i = 0; i < this.components.Length; i++)
                        {
                            if (!path.components[i].Equals(this.components[i]))
                            {
                                flag = false;
                                break;
                            }
                        }
                    }
                }
            }
            return flag;
        }
        /// <summary>
        /// get the specified component
        /// </summary>
        /// <param name="n">which component (0 ... length() - 1)</param>
        /// <returns>the nth component;</returns>
        public virtual string GetComponent(int n)
        {
            return this.components[n];
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            if (this.hashcode == 0)
            {
                for (int i = 0; i < this.components.Length; i++)
                {
                    this.hashcode += this.components[i].GetHashCode();
                }
            }
            return this.hashcode;
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            int length = this.Length;
            builder.Append(Path.DirectorySeparatorChar);
            for (int i = 0; i < length; i++)
            {
                builder.Append(this.GetComponent(i));
                if (i < (length - 1))
                {
                    builder.Append(Path.DirectorySeparatorChar);
                }
            }
            return builder.ToString();
        }

        /// <summary>
        /// Gets the length.
        /// </summary>
        /// <value>the number of components</value>
        public virtual int Length
        {
            get
            {
                return this.components.Length;
            }
        }
        /// <summary>
        /// Returns the path's parent or <c>null</c> if this path
        /// is the root path.
        /// </summary>
        /// <value>path of parent, or null if this path is the root path</value>
        public virtual POIFSDocumentPath Parent
        {
            get
            {
                int length = this.components.Length - 1;
                if (length < 0)
                {
                    return null;
                }
                POIFSDocumentPath path = new POIFSDocumentPath(null);
                path.components = new string[length];
                Array.Copy(this.components, 0, path.components, 0, length);
                return path;
            }
        }
    }
}
