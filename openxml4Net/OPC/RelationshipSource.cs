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
using System.Text;
using NPOI.OpenXml4Net.Exceptions;

namespace NPOI.OpenXml4Net.OPC
{
    interface RelationshipSource
    {
        /// <summary>
        /// Add a relationship to a part (except relationships part).
        /// </summary>
        /// <param name="targetPartName">
        /// Name of the target part. This one must be relative to the
        /// source root directory of the part.
        /// </param>
        /// <param name="targetMode">
        /// Mode [Internal|External].
        /// </param>
        /// <param name="relationshipType">
        /// Type of relationship.
        /// </param>
        /// <returns>The newly created and added relationship</returns>
        PackageRelationship AddRelationship(
                PackagePartName targetPartName, TargetMode targetMode,
                String relationshipType);

        /// <summary>
        /// <para>
        /// Add a relationship to a part (except relationships part).
        /// </para>
        /// <para>
        /// Check rule M1.25: The Relationships part shall not have relationships to
        /// any other part. Package implementers shall enforce this requirement upon
        /// the attempt to create such a relationship and shall treat any such
        /// relationship as invalid.
        /// </para>
        /// </summary>
        /// <param name="targetPartName">
        /// Name of the target part. This one must be relative to the
        /// source root directory of the part.
        /// </param>
        /// <param name="targetMode">
        /// Mode [Internal|External].
        /// </param>
        /// <param name="relationshipType">
        /// Type of relationship.
        /// </param>
        /// <param name="id">
        /// Relationship unique id.
        /// </param>
        /// <returns>The newly created and added relationship</returns>
        /// 
        /// <exception cref="InvalidFormatException">
        /// If the URI point to a relationship part URI.
        /// </exception>
        PackageRelationship AddRelationship(
                PackagePartName targetPartName, TargetMode targetMode,
                String relationshipType, String id);

        /// <summary>
        /// <para>
        /// Adds an external relationship to a part
        ///  (except relationships part).
        /// </para>
        /// <para>
        /// The targets of external relationships are not
        ///  subject to the same validity checks that internal
        ///  ones are, as the contents is potentially
        ///  any file, URL or similar.
        /// </para>
        /// </summary>
        /// <param name="target">External target of the relationship</param>
        /// <param name="relationshipType">Type of relationship.</param>
        /// <returns>The newly created and added relationship</returns>
        /// @see org.apache.poi.OpenXml4Net.opc.RelationshipSource#addExternalRelationship(java.lang.String, java.lang.String)
        PackageRelationship AddExternalRelationship(String target, String relationshipType);

        /// <summary>
        /// <para>
        /// Adds an external relationship to a part
        ///  (except relationships part).
        /// </para>
        /// <para>
        /// The targets of external relationships are not
        ///  subject to the same validity checks that internal
        ///  ones are, as the contents is potentially
        ///  any file, URL or similar.
        /// </para>
        /// </summary>
        /// <param name="target">External target of the relationship</param>
        /// <param name="relationshipType">Type of relationship.</param>
        /// <param name="id">Relationship unique id.</param>
        /// <returns>The newly created and added relationship</returns>
        /// @see org.apache.poi.OpenXml4Net.opc.RelationshipSource#addExternalRelationship(java.lang.String, java.lang.String)
        PackageRelationship AddExternalRelationship(String target, String relationshipType, String id);

        /// <summary>
        /// Delete all the relationships attached to this.
        /// </summary>
        void ClearRelationships();

        /// <summary>
        /// Delete the relationship specified by its id.
        /// </summary>
        /// <param name="id">
        /// The ID identified the part to delete.
        /// </param>
        void RemoveRelationship(String id);

        /// <summary>
        /// Retrieve all the relationships attached to this.
        /// </summary>
        /// <returns>This part's relationships.</returns>
        /// <exception cref="OpenXml4NetException"></exception>
        PackageRelationshipCollection Relationships { get; }

        /// <summary>
        /// Retrieves a package relationship from its id.
        /// </summary>
        /// <param name="id">
        /// ID of the package relationship to retrieve.
        /// </param>
        /// <returns>The package relationship</returns>
        PackageRelationship GetRelationship(String id);

        /// <summary>
        /// Retrieve all relationships attached to this part which have the specified
        /// type.
        /// </summary>
        /// <param name="relationshipType">
        /// Relationship type filter.
        /// </param>
        /// <returns>All relationships from this part that have the specified type.</returns>
        /// <exception cref="InvalidFormatException">
        /// If an error occurs while parsing the part.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// If the package is open in write only mode.
        /// </exception>
        PackageRelationshipCollection GetRelationshipsByType(
               String relationshipType);

        /// <summary>
        /// Knows if the part have any relationships.
        /// </summary>
        /// <returns><b>true</b> if the part have at least one relationship else
        /// <b>false</b>.
        /// </returns>
        bool HasRelationships { get; }

        /// <summary>
        /// Checks if the specified relationship is part of this package part.
        /// </summary>
        /// <param name="rel">
        /// The relationship to check.
        /// </param>
        /// <returns><b>true</b> if the specified relationship exists in this part,
        /// else returns <b>false</b>
        /// </returns>
        bool IsRelationshipExists(PackageRelationship rel);

    }
}
