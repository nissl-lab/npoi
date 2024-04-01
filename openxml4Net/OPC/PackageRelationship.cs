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

using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{
    /// <summary>
    /// A part relationship.
    /// 
    /// @author Julien Chable
    /// @version 1.0
    /// 
    /// </summary>
    public class PackageRelationship
    {

        private static Uri containerRelationshipPart = PackagingUriHelper.ParseUri("/_rels/.rels", UriKind.RelativeOrAbsolute);

        /* XML markup */

        public static String ID_ATTRIBUTE_NAME = "Id";

        public static String RELATIONSHIPS_TAG_NAME = "Relationships";

        public static String RELATIONSHIP_TAG_NAME = "Relationship";

        public static String TARGET_ATTRIBUTE_NAME = "Target";

        public static String TARGET_MODE_ATTRIBUTE_NAME = "TargetMode";

        public static String TYPE_ATTRIBUTE_NAME = "Type";

        /* End XML markup */

        /// <summary>
        /// Relation id.
        /// </summary>
        private String id;

        /// <summary>
        /// Reference to the package.
        /// </summary>
        private OPCPackage container;

        /// <summary>
        /// Relationship type
        /// </summary>
        private String relationshipType;

        /// <summary>
        /// Part of this relationship source
        /// </summary>
        private PackagePart source;

        /// <summary>
        /// Targeting mode [Internal|External]
        /// </summary>
        private TargetMode? targetMode;

        /// <summary>
        /// Target URI
        /// </summary>
        private Uri targetUri;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="pkg"></param>
        /// <param name="sourcePart"></param>
        /// <param name="targetUri"></param>
        /// <param name="targetMode"></param>
        /// <param name="relationshipType"></param>
        /// <param name="id"></param>
        public PackageRelationship(OPCPackage pkg, PackagePart sourcePart,
                Uri targetUri, TargetMode targetMode, String relationshipType,
                String id)
        {
            if(pkg == null)
                throw new ArgumentException("pkg");
            if(targetUri == null)
                throw new ArgumentException("targetUri");
            if(relationshipType == null)
                throw new ArgumentException("relationshipType");
            if(id == null)
                throw new ArgumentException("id");

            this.container = pkg;
            this.source = sourcePart;
            this.targetUri = targetUri;
            this.targetMode = targetMode;
            this.relationshipType = relationshipType;
            this.id = id;
        }


        public override bool Equals(Object obj)
        {
            if(!(obj is PackageRelationship))
            {
                return false;
            }
            PackageRelationship rel = (PackageRelationship)obj;
            return (this.id == rel.id
                    && this.relationshipType == rel.relationshipType
                    && (rel.source != null ? rel.source.Equals(this.source) : true)
                    && this.targetMode == rel.targetMode && this.targetUri
                    .Equals(rel.targetUri));
        }


        public override int GetHashCode()
        {
            return this.id.GetHashCode() + this.relationshipType.GetHashCode()
                + this.source?.GetHashCode() ?? 0 + this.targetMode.GetHashCode()
                + this.targetUri.GetHashCode();
        }

        /* Getters */

        public static Uri ContainerPartRelationship
        {
            get
            {
                return containerRelationshipPart;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>the container</returns>
        public OPCPackage Package
        {
            get
            {
                return container;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>the id</returns>
        public String Id
        {
            get
            {
                return id;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>the relationshipType</returns>
        public String RelationshipType
        {
            get
            {
                return relationshipType;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>the source</returns>
        public PackagePart Source
        {
            get
            {
                return source;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>URL of the source part of this relationship</returns>
        public Uri SourceUri
        {
            get
            {
                if(source == null)
                {
                    return PackagingUriHelper.PACKAGE_ROOT_URI;
                }
                return source.PartName.URI;
            }
        }

        /// <summary>
        /// public URI getSourceUri(){ }
        /// </summary>
        /// <returns>the targetMode</returns>
        public TargetMode? TargetMode
        {
            get
            {
                return targetMode;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>the targetUri</returns>
        public Uri TargetUri
        {
            get
            {
                // If it's an external target, we don't
                //  need to apply our normal validation rules
                if(targetMode == OPC.TargetMode.External)
                {
                    return targetUri;
                }

                // Internal target
                // If it isn't absolute, resolve it relative
                //  to ourselves
                if(!targetUri.ToString().StartsWith("/"))
                {
                    // So it's a relative part name, try to resolve it
                    return PackagingUriHelper.ResolvePartUri(SourceUri, targetUri);
                }
                return targetUri;
            }
        }


        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(id == null ? "id=null" : "id=" + id);
            sb.Append(container == null ? " - container=null" : " - container="
                    + container.ToString());
            sb.Append(relationshipType == null ? " - relationshipType=null"
                    : " - relationshipType=" + relationshipType.ToString());
            sb.Append(source == null ? " - source=null" : " - source="
                    + SourceUri.OriginalString);
            sb.Append(targetUri == null ? " - target=null" : " - target="
                    + TargetUri.OriginalString);
            sb.Append(targetMode == null ? ",targetMode=null" : ",targetMode="
                    + targetMode.ToString());
            return sb.ToString();
        }
    }
}