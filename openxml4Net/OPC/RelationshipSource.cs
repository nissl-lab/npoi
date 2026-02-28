using System;
using System.Text;
using NPOI.OpenXml4Net.Exceptions;

namespace NPOI.OpenXml4Net.OPC
{
     interface RelationshipSource
    {
        /**
 * Add a relationship to a part (except relationships part).
 * 
 * @param targetPartName
 *            Name of the target part. This one must be relative to the
 *            source root directory of the part.
 * @param targetMode
 *            Mode [Internal|External].
 * @param relationshipType
 *            Type of relationship.
 * @return The newly created and added relationship
 */
        PackageRelationship AddRelationship(
                PackagePartName targetPartName, TargetMode targetMode,
                String relationshipType);

        /**
         * Add a relationship to a part (except relationships part).
         
         * Check rule M1.25: The Relationships part shall not have relationships to
         * any other part. Package implementers shall enforce this requirement upon
         * the attempt to create such a relationship and shall treat any such
         * relationship as invalid.
         * 
         * @param targetPartName
         *            Name of the target part. This one must be relative to the
         *            source root directory of the part.
         * @param targetMode
         *            Mode [Internal|External].
         * @param relationshipType
         *            Type of relationship.
         * @param id
         *            Relationship unique id.
         * @return The newly created and added relationship
         * 
         * @throws InvalidFormatException
         *             If the URI point to a relationship part URI.
         */
        PackageRelationship AddRelationship(
                PackagePartName targetPartName, TargetMode targetMode,
                String relationshipType, String id);

        /**
         * Adds an external relationship to a part
         *  (except relationships part).
         * 
         * The targets of external relationships are not
         *  subject to the same validity checks that internal
         *  ones are, as the contents is potentially
         *  any file, URL or similar.
         *  
         * @param target External target of the relationship
         * @param relationshipType Type of relationship.
         * @return The newly created and added relationship
         * @see org.apache.poi.OpenXml4Net.opc.RelationshipSource#addExternalRelationship(java.lang.String, java.lang.String)
         */
        PackageRelationship AddExternalRelationship(String target, String relationshipType);

        /**
         * Adds an external relationship to a part
         *  (except relationships part).
         * 
         * The targets of external relationships are not
         *  subject to the same validity checks that internal
         *  ones are, as the contents is potentially
         *  any file, URL or similar.
         *  
         * @param target External target of the relationship
         * @param relationshipType Type of relationship.
         * @param id Relationship unique id.
         * @return The newly created and added relationship
         * @see org.apache.poi.OpenXml4Net.opc.RelationshipSource#addExternalRelationship(java.lang.String, java.lang.String)
         */
        PackageRelationship AddExternalRelationship(String target, String relationshipType, String id);

        /**
         * Delete all the relationships attached to this.
         */
        void ClearRelationships();

        /**
         * Delete the relationship specified by its id.
         * 
         * @param id
         *            The ID identified the part to delete.
         */
        void RemoveRelationship(String id);

        /**
         * Retrieve all the relationships attached to this.
         * 
         * @return This part's relationships.
         * @throws OpenXml4NetException
         */
        PackageRelationshipCollection Relationships { get; }

        /**
         * Retrieves a package relationship from its id.
         * 
         * @param id
         *            ID of the package relationship to retrieve.
         * @return The package relationship
         */
         PackageRelationship GetRelationship(String id);

        /**
         * Retrieve all relationships attached to this part which have the specified
         * type.
         * 
         * @param relationshipType
         *            Relationship type filter.
         * @return All relationships from this part that have the specified type.
         * @throws InvalidFormatException
         *             If an error occurs while parsing the part.
         * @throws InvalidOperationException
         *             If the package is open in write only mode.
         */
         PackageRelationshipCollection GetRelationshipsByType(
                String relationshipType);

        /**
         * Knows if the part have any relationships.
         * 
         * @return <b>true</b> if the part have at least one relationship else
         *         <b>false</b>.
         */
         bool HasRelationships { get; }

        /**
         * Checks if the specified relationship is part of this package part.
         * 
         * @param rel
         *            The relationship to check.
         * @return <b>true</b> if the specified relationship exists in this part,
         *         else returns <b>false</b>
         */
         bool IsRelationshipExists(PackageRelationship rel);

    }
}
