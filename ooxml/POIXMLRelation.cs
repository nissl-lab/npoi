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
using NPOI.OpenXml4Net.OPC;
using System;
using System.Text.RegularExpressions;

namespace NPOI
{

    /**
     * Represents a descriptor of a OOXML relation.
     *
     * @author Yegor Kozlov
     */
    public abstract class POIXMLRelation
    {

        /**
         * Describes the content stored in a part.
         */
        private String _type;

        /**
         * The kind of connection between a source part and a target part in a namespace.
         */
        private String _relation;

        /**
         * The path component of a pack URI.
         */
        private String _defaultName;

        /**
         * Defines what object is used to construct instances of this relationship
         */
        private Type _cls;

        /**
         * Defines a function to construct an instance of this relationship with parent and part provided
         */
        private Func<POIXMLDocumentPart, PackagePart, POIXMLDocumentPart> _createPartWithParent;

        /**
         * Defines a function to construct an instance of this relationship with part provided
         */
        private Func<PackagePart, POIXMLDocumentPart> _createPart;

        /**
         * Defines a function to construct an instance of this relationship
         */
        private Func<POIXMLDocumentPart> _createInstance;

        /**
         * Instantiates a POIXMLRelation.
         *
         * @param type content type
         * @param rel  relationship
         * @param defaultName default item name
         * @param cls defines what object is used to construct instances of this relationship
         */
        public POIXMLRelation(String type, String rel, String defaultName, Type cls
            , Func<POIXMLDocumentPart, PackagePart, POIXMLDocumentPart> createPartWithParent
            , Func<PackagePart, POIXMLDocumentPart> createPart
            , Func<POIXMLDocumentPart> createInstance)
        {
            _type = type;
            _relation = rel;
            _defaultName = defaultName;
            _cls = cls;
            _createPartWithParent=createPartWithParent;
            _createPart=createPart;
            _createInstance=createInstance;
        }

        /**
         * Instantiates a POIXMLRelation.
         *
         * @param type content type
         * @param rel  relationship
         * @param defaultName default item name
         */
        public POIXMLRelation(String type, String rel, String defaultName)
            : this(type, rel, defaultName, null, null, null, null)
        {

        }
        /**
         * Return the content type. Content types define a media type, a subtype, and an
         * optional set of parameters, as defined in RFC 2616.
         *
         * @return the content type
         */
        public String ContentType
        {
            get
            {
                return _type;
            }
        }

        /**
         * Return the relationship, the kind of connection between a source part and a target part in a namespace.
         * Relationships make the connections between parts directly discoverable without looking at the content
         * in the parts, and without altering the parts themselves.
         *
         * @return the relationship
         */
        public String Relation
        {
            get
            {
                return _relation;
            }
        }

        /**
         * Return the default part name. Part names are used to refer to a part in the context of a
         * namespace, typically as part of a URI.
         *
         * @return the default part name
         */
        public String DefaultFileName
        {
            get
            {
                return _defaultName;
            }
        }

        /**
         * Returns the filename for the nth one of these,
         *  e.g. /xl/comments4.xml
         */
        public String GetFileName(int index)
        {
            if (_defaultName.IndexOf("#") == -1)
            {
                // Generic filename in all cases
                return DefaultFileName;
            }
            return _defaultName.Replace("#", index.ToString());
        }

        /**
         * Returns the index of the filename within the package for the given part.
         *  e.g. 4 for /xl/comments4.xml
         */
        public int GetFileNameIndex(POIXMLDocumentPart part)
        {
            Regex regex = new Regex(_defaultName.Replace("#", "(\\d+)"));
            return int.Parse(regex.Match(part.GetPackagePart().PartName.Name).Groups[1].Value);
            //return Integer.valueOf(part.getPackagePart().getPartName().getName().replaceAll(regex, "$1"));
        }
        /**
         * Return type of the object used to construct instances of this relationship
         *
         * @return the class of the object used to construct instances of this relation
         */
        public Type RelationClass
        {
            get
            {
                return _cls;
            }
        }

        /**
         * Function to construct an instance of this relation type with parent and package part provided
         *
         * @return An instance of the document part of this relation
         */
        public Func<POIXMLDocumentPart, PackagePart, POIXMLDocumentPart> CreatePartWithParent
        {
            get
            {
                return _createPartWithParent;
            }
        }

        /**
         * Function to construct an instance of this relation type with package part provided
         *
         * @return An instance of the document part of this relation
         */
        public Func<PackagePart, POIXMLDocumentPart> CreatePart
        {
            get
            {
                return _createPart;
            }
        }

        /**
         * Function to construct an instance of this relation type
         *
         * @return An instance of the document part of this relation
         */
        public Func<POIXMLDocumentPart> CreateInstance
        {
            get
            {
                return _createInstance;
            }
        }
    }
}






