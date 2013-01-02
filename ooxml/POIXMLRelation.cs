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
        protected String _type;

        /**
         * The kind of connection between a source part and a target part in a namespace.
         */
        protected String _relation;

        /**
         * The path component of a pack URI.
         */
        protected String _defaultName;

        /**
         * Defines what object is used to construct instances of this relationship
         */
        private Type _cls;

        /**
         * Instantiates a POIXMLRelation.
         *
         * @param type content type
         * @param rel  relationship
         * @param defaultName default item name
         * @param cls defines what object is used to construct instances of this relationship
         */
        public POIXMLRelation(String type, String rel, String defaultName, Type cls)
        {
            _type = type;
            _relation = rel;
            _defaultName = defaultName;
            _cls = cls;
        }

        /**
         * Instantiates a POIXMLRelation.
         *
         * @param type content type
         * @param rel  relationship
         * @param defaultName default item name
         */
        public POIXMLRelation(String type, String rel, String defaultName)
            : this(type, rel, defaultName, null)
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
         * Return type of the obejct used to construct instances of this relationship
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
    }
}






