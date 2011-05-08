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
using System.Collections;
using System.IO;

namespace NPOI.POIFS.Properties
{
    /// <summary>
    /// Trivial extension of Property for POIFSDocuments
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// </summary>
    public class DirectoryProperty:Property,Parent
    {
        // List of Property instances
        private IList _children;

        // Set of children's names
        private IList  _children_names;

        public override void Dispose()
        {
            _children = null;
            _children_names = null;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DirectoryProperty"/> class.
        /// </summary>
        /// <param name="name">the name of the directory</param>
        public DirectoryProperty(String name):base()
        {
            _children       = new ArrayList();
            _children_names = new ArrayList();
            Name=name;
            Size=0;
            PropertyType=PropertyConstants.DIRECTORY_TYPE;
            StartBlock=0;
            NodeColor=_NODE_BLACK;   // simplification
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DirectoryProperty"/> class.
        /// </summary>
        /// <param name="index">index number</param>
        /// <param name="array">byte data</param>
        /// <param name="offset">offset into byte data</param>
        public DirectoryProperty(int index, byte [] array,int offset):base(index, array, offset)
        {
            _children       = new ArrayList();
            _children_names = new ArrayList();
        }


        /// <summary>
        /// Change a Property's name
        /// </summary>
        /// <param name="property">the Property whose name Is being Changed.</param>
        /// <param name="newName">the new name for the Property</param>
        /// <returns>true if the name Change could be made, else false</returns>
        public bool ChangeName(Property property, String newName)
        {
            bool result;
            String  oldName = property.Name;

            property.Name=newName;
            String cleanNewName = property.Name;

            if (_children_names.Contains(cleanNewName))
            {

                // revert the Change
                property.Name=oldName;
                result = false;
            }
            else
            {
                _children_names.Add(cleanNewName);
                _children_names.Remove(oldName);
                result = true;
            }
            return result;
        }

        /// <summary>
        /// Delete a Property
        /// </summary>
        /// <param name="property">the Property being Deleted</param>
        /// <returns>true if the Property could be Deleted, else false</returns>
        public bool DeleteChild(Property property)
        {

            bool flag = this._children.Contains(property);
            this._children.Remove(property);
            bool flag2 = flag;
            if (flag2)
            {
                this._children_names.Remove(property.Name);
            }
            return flag2;

        }

        /// <summary>
        /// Directory Property Comparer
        /// </summary>
        public class PropertyComparator : IComparer
        {

            /// <summary>
            /// Object equality, implemented as object identity
            /// </summary>
            /// <param name="o">Object we're being Compared to</param>
            /// <returns>true if identical, else false</returns>
            public override bool Equals(Object o)
            {
                return this == o;
            }


            /// <summary>
            /// Compare method. Assumes both parameters are non-null
            /// instances of Property. One property Is less than another if
            /// its name Is shorter than the other property's name. If the
            /// names are the same length, the property whose name comes
            /// before the other property's name, alphabetically, Is less
            /// than the other property.
            /// </summary>
            /// <param name="o1">first object to Compare, better be a Property</param>
            /// <param name="o2">second object to Compare, better be a Property</param>
            /// <returns>negative value if o1 smaller than o2,
            ///         zero           if o1 equals o2,
            ///        positive value if o1 bigger than  o2.</returns>
            public int Compare(Object o1, Object o2)
            {
                //solution from http://mail-archives.apache.org/mod_mbox/poi-dev/200604.mbox/%3Cbug-39234-7501@http.issues.apache.org/bugzilla/%3E

                String VBA_PROJECT = "_VBA_PROJECT";

                String name1 = ((Property)o1).Name;
                String name2 = ((Property)o2).Name;

                int num = name1.Length - name2.Length;

                if (num == 0)
                {
                    if (name1.CompareTo(VBA_PROJECT) == 0)
                    {
                        num = 1;
                    }
                    else if (name2.CompareTo(VBA_PROJECT) == 0)
                    {
                        num = -1;
                    }
                    else
                    {
                        if (name1.StartsWith("__") && name2.StartsWith("__"))
                        {
                            // Betweeen __SRP_0 and __SRP_1 just sort as normal
                            num = String.Compare(name1, name2, true);
                        }
                        else if (name1.StartsWith("__"))
                        {
                            // If only name1 is __XXX then this will be placed after name2
                            num = 1;
                        }
                        else if (name2.StartsWith("__"))
                        {
                            // If only name2 is __XXX then this will be placed after name1
                            num = -1;
                        }
                        else
                        {
                            // result = name1.compareTo(name2);
                            // The default case is to sort names ignoring case
                            num = String.Compare(name1, name2, true);
                        }
                    }

                }
                return num;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is directory.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if a directory type Property; otherwise, <c>false</c>.
        /// </value>
        public override bool IsDirectory
        {
            get{return true;}
        }

        /// <summary>
        /// Perform whatever activities need to be performed prior to
        /// writing
        /// </summary>
        public override void PreWrite()
        {
            if (_children.Count > 0)
            {
                Property[] children = new Property[this._children.Count];

                this._children.CopyTo(children, 0);
                Array.Sort(children, new PropertyComparator());
                int midpoint = children.Length / 2;

                this.ChildProperty=children[ midpoint ].Index;
                children[ 0 ].PreviousChild=null;
                children[ 0 ].NextChild=null;
                for (int j = 1; j < midpoint; j++)
                {
                    children[ j ].PreviousChild=children[ j - 1 ];
                    children[ j ].NextChild=null;
                }
                if (midpoint != 0)
                {
                    children[ midpoint ]
                        .PreviousChild=children[ midpoint - 1 ];
                }
                if (midpoint != (children.Length - 1))
                {
                    children[ midpoint ].NextChild=children[ midpoint + 1 ];
                    for (int j = midpoint + 1; j < children.Length - 1; j++)
                    {
                        children[ j ].PreviousChild=null;
                        children[ j ].NextChild=children[ j + 1 ];
                    }
                    children[ children.Length - 1 ].PreviousChild=null;
                    children[ children.Length - 1 ].NextChild=null;
                }
                else
                {
                    children[ midpoint ].NextChild=null;
                }
            }
        }

        /// <summary>
        /// Get an iterator over the children of this Parent; all elements
        /// are instances of Property.
        /// </summary>
        /// <value>Iterator of children; may refer to an empty collection</value>
        public IEnumerator Children
        {
            get{return _children.GetEnumerator();}
        }

        /// <summary>
        /// Add a new child to the collection of children
        /// </summary>
        /// <param name="property">the new child to be added; must not be null</param>
        public void AddChild(Property property)
        {
            String name = property.Name;

            if (_children_names.Contains(name))
            {
                throw new IOException("Duplicate name \"" + name + "\"");
            }
            _children_names.Add(name);
            _children.Add(property);
        }
    }
}
