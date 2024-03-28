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
using System.Text.RegularExpressions;

namespace NPOI
{

    /// <summary>
    /// Represents a descriptor of a OOXML relation.
    /// </summary>
    /// @author Yegor Kozlov
    public abstract class POIXMLRelation
    {

        /// <summary>
        /// Describes the content stored in a part.
        /// </summary>
        private String _type;

        /// <summary>
        /// The kind of connection between a source part and a target part in a namespace.
        /// </summary>
        private String _relation;

        /// <summary>
        /// The path component of a pack URI.
        /// </summary>
        private String _defaultName;

        /// <summary>
        /// Defines what object is used to construct instances of this relationship
        /// </summary>
        private Type _cls;

        /// <summary>
        /// Instantiates a POIXMLRelation.
        /// </summary>
        /// <param name="type">content type</param>
        /// <param name="rel"> relationship</param>
        /// <param name="defaultName">default item name</param>
        /// <param name="cls">defines what object is used to construct instances of this relationship</param>
        public POIXMLRelation(String type, String rel, String defaultName, Type cls)
        {
            _type = type;
            _relation = rel;
            _defaultName = defaultName;
            _cls = cls;
        }

        /// <summary>
        /// Instantiates a POIXMLRelation.
        /// </summary>
        /// <param name="type">content type</param>
        /// <param name="rel"> relationship</param>
        /// <param name="defaultName">default item name</param>
        public POIXMLRelation(String type, String rel, String defaultName)
            : this(type, rel, defaultName, null)
        {

        }
        /// <summary>
        /// Return the content type. Content types define a media type, a subtype, and an
        /// optional set of parameters, as defined in RFC 2616.
        /// </summary>
        /// <returns>the content type</returns>
        public String ContentType
        {
            get
            {
                return _type;
            }
        }

        /// <summary>
        /// Return the relationship, the kind of connection between a source part and a target part in a namespace.
        /// Relationships make the connections between parts directly discoverable without looking at the content
        /// in the parts, and without altering the parts themselves.
        /// </summary>
        /// <returns>the relationship</returns>
        public String Relation
        {
            get
            {
                return _relation;
            }
        }

        /// <summary>
        /// Return the default part name. Part names are used to refer to a part in the context of a
        /// namespace, typically as part of a URI.
        /// </summary>
        /// <returns>the default part name</returns>
        public String DefaultFileName
        {
            get
            {
                return _defaultName;
            }
        }

        /// <summary>
        /// Returns the filename for the nth one of these,
        ///  e.g. /xl/comments4.xml
        /// </summary>
        public String GetFileName(int index)
        {
            if(_defaultName.IndexOf("#") == -1)
            {
                // Generic filename in all cases
                return DefaultFileName;
            }
            return _defaultName.Replace("#", index.ToString());
        }

        /// <summary>
        /// Returns the index of the filename within the package for the given part.
        ///  e.g. 4 for /xl/comments4.xml
        /// </summary>
        public int GetFileNameIndex(POIXMLDocumentPart part)
        {
            Regex regex = new Regex(_defaultName.Replace("#", "(\\d+)"));
            return int.Parse(regex.Match(part.GetPackagePart().PartName.Name).Groups[1].Value);
            //return Integer.valueOf(part.getPackagePart().getPartName().getName().replaceAll(regex, "$1"));
        }
        /// <summary>
        /// Return type of the obejct used to construct instances of this relationship
        /// </summary>
        /// <returns>the class of the object used to construct instances of this relation</returns>
        public Type RelationClass
        {
            get
            {
                return _cls;
            }
        }
    }
}






