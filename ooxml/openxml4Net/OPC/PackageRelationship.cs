using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{
    /**
* A part relationship.
* 
* @author Julien Chable
* @version 1.0
*/
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

        /**
         * L'ID de la relation.
         */
        private String id;

        /**
         * Reference to the package.
         */
        private OPCPackage container;

        /**
         * Type de relation.
         */
        private String relationshipType;

        /**
         * Partie source de cette relation.
         */
        private PackagePart source;

        /**
         * Le mode de ciblage [Internal|External]
         */
        private TargetMode? targetMode;

        /**
         * URI de la partie cible.
         */
        private Uri targetUri;

        /**
         * Constructor.
         * 
         * @param pkg
         * @param sourcePart
         * @param targetUri
         * @param targetMode
         * @param relationshipType
         * @param id
         */
        public PackageRelationship(OPCPackage pkg, PackagePart sourcePart,
                Uri targetUri, TargetMode targetMode, String relationshipType,
                String id)
        {
            if (pkg == null)
                throw new ArgumentException("pkg");
            if (targetUri == null)
                throw new ArgumentException("targetUri");
            if (relationshipType == null)
                throw new ArgumentException("relationshipType");
            if (id == null)
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
            if (!(obj is PackageRelationship))
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
                    + this.source.GetHashCode() + this.targetMode.GetHashCode()
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

        /**
         * @return the container
         */
        public OPCPackage Package
        {
            get
            {
                return container;
            }
        }

        /**
         * @return the id
         */
        public String Id
        {
            get
            {
                return id;
            }
        }

        /**
         * @return the relationshipType
         */
        public String RelationshipType
        {
            get
            {
                return relationshipType;
            }
        }

        /**
         * @return the source
         */
        public PackagePart Source
        {
            get
            {
                return source;
            }
        }

        /**
         * 
         * @return URL of the source part of this relationship
         */
        public Uri SourceUri
        {
            get
            {
                if (source == null)
                {
                    return PackagingUriHelper.PACKAGE_ROOT_URI;
                }
                return source.PartName.URI;
            }
        }

        /**
         * public URI getSourceUri(){ }
         * 
         * @return the targetMode
         */
        public TargetMode? TargetMode
        {
            get
            {
                return targetMode;
            }
        }

        /**
         * @return the targetUri
         */
        public Uri TargetUri
        {
            get
            {
                // If it's an external target, we don't
                //  need to apply our normal validation rules
                if (targetMode == NPOI.OpenXml4Net.OPC.TargetMode.External)
                {
                    return targetUri;
                }

                // Internal target
                // If it isn't absolute, resolve it relative
                //  to ourselves
                if (!targetUri.ToString().StartsWith("/"))
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