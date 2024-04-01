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
using System.IO;
using NPOI.OpenXml4Net.OPC.Internal;
using NPOI.OpenXml4Net.Exceptions;


namespace NPOI.OpenXml4Net.OPC
{
    /// <summary>
    /// Provides a base class for parts stored in a Package.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 0.9
    /// </remarks>

    public abstract class PackagePart : RelationshipSource, IComparable<PackagePart>
    {

        /// <summary>
        /// This part's container.
        /// </summary>
        internal OPCPackage _container;

        /// <summary>
        /// The part name. (required by the specification [M1.1])
        /// </summary>
        protected PackagePartName _partName;

        /// <summary>
        /// The type of content of this part. (required by the specification [M1.2])
        /// </summary>
        internal ContentType _contentType;

        /// <summary>
        /// Flag to know if this part is a relationship.
        /// </summary>
        private bool _isRelationshipPart;

        /// <summary>
        /// Flag to know if this part has been logically deleted.
        /// </summary>
        private bool _isDeleted;

        /// <summary>
        /// This part's relationships.
        /// </summary>
        private PackageRelationshipCollection _relationships;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="pack">
        /// Parent package.
        /// </param>
        /// <param name="partName">
        /// The part name, relative to the parent Package root.
        /// </param>
        /// <param name="contentType">
        /// The content type.
        /// </param>
        /// <exception cref="InvalidFormatException">
        /// If the specified URI is not valid.
        /// </exception>
        protected PackagePart(OPCPackage pack, PackagePartName partName,
                ContentType contentType)
            : this(pack, partName, contentType, true)
        {

        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="pack">
        /// Parent package.
        /// </param>
        /// <param name="partName">
        /// The part name, relative to the parent Package root.
        /// </param>
        /// <param name="contentType">
        /// The content type.
        /// </param>
        /// <param name="loadRelationships">
        /// Specify if the relationships will be loaded
        /// </param>
        /// <exception cref="InvalidFormatException">
        /// If the specified URI is not valid.
        /// </exception>
        protected PackagePart(OPCPackage pack, PackagePartName partName,
                ContentType contentType, bool loadRelationships)
        {
            this._partName = partName;
            this._contentType = contentType;
            this._container = (ZipPackage)pack; // TODO - enforcing ZipPackage here - perhaps should change constructor signature

            // Check if this part is a relationship part
            _isRelationshipPart = this._partName.IsRelationshipPartURI();

            // Load relationships if any
            if (loadRelationships)
                LoadRelationships();
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="pack">
        /// Parent package.
        /// </param>
        /// <param name="partName">
        /// The part name, relative to the parent Package root.
        /// </param>
        /// <param name="contentType">
        /// The Multipurpose Internet Mail Extensions (MIME) content type
        /// of the part's data stream.
        /// </param>
        public PackagePart(OPCPackage pack, PackagePartName partName,
                String contentType)
            : this(pack, partName, new ContentType(contentType))
        {

        }
        /// <summary>
        /// Check if the new part was already added before via PackagePart.addRelationship()
        /// </summary>
        /// <param name="packagePart">to find the relationship for</param>
        /// <returns>The existing relationship, or null if there isn't yet one</returns>
        public PackageRelationship FindExistingRelation(PackagePart packagePart)
        {
            return _relationships.FindExistingInternalRelation(packagePart);
        }
        /// <summary>
        /// <para>
        /// Adds an external relationship to a part (except relationships part).
        /// </para>
        /// <para>
        /// The targets of external relationships are not subject to the same
        /// validity checks that internal ones are, as the contents is potentially
        /// any file, URL or similar.
        /// </para>
        /// </summary>
        /// <param name="target">
        /// External target of the relationship
        /// </param>
        /// <param name="relationshipType">
        /// Type of relationship.
        /// </param>
        /// <returns>The newly created and added relationship</returns>
        /// @see OPC.RelationshipSource#addExternalRelationship(java.lang.String,
        ///      java.lang.String)
        public PackageRelationship AddExternalRelationship(String target,
                String relationshipType)
        {
            return AddExternalRelationship(target, relationshipType, null);
        }

        /// <summary>
        /// <para>
        /// Adds an external relationship to a part (except relationships part).
        /// </para>
        /// <para>
        /// The targets of external relationships are not subject to the same
        /// validity checks that internal ones are, as the contents is potentially
        /// any file, URL or similar.
        /// </para>
        /// </summary>
        /// <param name="target">
        /// External target of the relationship
        /// </param>
        /// <param name="relationshipType">
        /// Type of relationship.
        /// </param>
        /// <param name="id">
        /// Relationship unique id.
        /// </param>
        /// <returns>The newly created and added relationship</returns>
        /// @see OPC.RelationshipSource#addExternalRelationship(java.lang.String,
        ///      java.lang.String)
        public PackageRelationship AddExternalRelationship(String target,
                String relationshipType, String id)
        {
            if (target == null)
            {
                throw new ArgumentException("target");
            }
            if (relationshipType == null)
            {
                throw new ArgumentException("relationshipType");
            }

            if (_relationships == null)
            {
                _relationships = new PackageRelationshipCollection();
            }

            Uri targetURI;
            try
            {
                targetURI = PackagingUriHelper.ParseUri(target,UriKind.RelativeOrAbsolute);
            }
            catch (UriFormatException e)
            {
                throw new ArgumentException("Invalid target - " + e);
            }

            return _relationships.AddRelationship(targetURI, TargetMode.External,
                    relationshipType, id);
        }

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
        /// @see OPC.RelationshipSource#AddRelationship(OPC.PackagePartName,
        ///      OPC.TargetMode, java.lang.String)
        public PackageRelationship AddRelationship(PackagePartName targetPartName,
                TargetMode targetMode, String relationshipType)
        {
            return AddRelationship(targetPartName, targetMode, relationshipType,
                    null);
        }

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
        /// @see OPC.RelationshipSource#AddRelationship(OPC.PackagePartName,
        ///      OPC.TargetMode, java.lang.String, java.lang.String)
        public PackageRelationship AddRelationship(PackagePartName targetPartName,
                TargetMode targetMode, String relationshipType, String id)
        {
            _container.ThrowExceptionIfReadOnly();

            if (targetPartName == null)
            {
                throw new ArgumentException("targetPartName");
            }
            //if (targetMode == null)
            //{
            //    throw new ArgumentException("targetMode");
            //}
            if (relationshipType == null)
            {
                throw new ArgumentException("relationshipType");
            }

            if (this.IsRelationshipPart || targetPartName.IsRelationshipPartURI())
            {
                throw new InvalidOperationException(
                        "Rule M1.25: The Relationships part shall not have relationships to any other part.");
            }

            if (_relationships == null)
            {
                _relationships = new PackageRelationshipCollection();
            }

            return _relationships.AddRelationship(targetPartName.URI,
                    targetMode, relationshipType, id);
        }

        /// <summary>
        /// Add a relationship to a part (except relationships part).
        /// </summary>
        /// <param name="targetURI">
        /// URI the target part. Must be relative to the source root
        /// directory of the part.
        /// </param>
        /// <param name="targetMode">
        /// Mode [Internal|External].
        /// </param>
        /// <param name="relationshipType">
        /// Type of relationship.
        /// </param>
        /// <returns>The newly created and added relationship</returns>
        /// @see OPC.RelationshipSource#AddRelationship(OPC.PackagePartName,
        ///      OPC.TargetMode, java.lang.String)
        public PackageRelationship AddRelationship(Uri targetURI,
                TargetMode targetMode, String relationshipType)
        {
            return AddRelationship(targetURI, targetMode, relationshipType, null);
        }

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
        /// <param name="targetURI">
        /// URI of the target part. Must be relative to the source root
        /// directory of the part.
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
        /// @see OPC.RelationshipSource#AddRelationship(OPC.PackagePartName,
        ///      OPC.TargetMode, java.lang.String, java.lang.String)
        public PackageRelationship AddRelationship(Uri targetURI,
                TargetMode targetMode, String relationshipType, String id)
        {
            _container.ThrowExceptionIfReadOnly();

            if (targetURI == null)
            {
                throw new ArgumentException("targetPartName");
            }
            //if (targetMode == null)
            //{
            //    throw new ArgumentException("targetMode");
            //}
            if (relationshipType == null)
            {
                throw new ArgumentException("relationshipType");
            }

            // Try to retrieve the target part

            if (this.IsRelationshipPart
                    || PackagingUriHelper.IsRelationshipPartURI(targetURI))
            {
                throw new InvalidOperationException(
                        "Rule M1.25: The Relationships part shall not have relationships to any other part.");
            }

            if (_relationships == null)
            {
                _relationships = new PackageRelationshipCollection();
            }

            return _relationships.AddRelationship(targetURI,
                    targetMode, relationshipType, id);
        }

        /// <summary>
        /// <see cref="OPC.RelationshipSource.clearRelationships()" />
        /// </summary>
        public void ClearRelationships()
        {
            if (_relationships != null)
            {
                _relationships.Clear();
            }
        }

        /// <summary>
        /// Delete the relationship specified by its id.
        /// </summary>
        /// <param name="id">
        /// The ID identified the part to delete.
        /// </param>
        /// @see OPC.RelationshipSource#removeRelationship(java.lang.String)
        public void RemoveRelationship(String id)
        {
            this._container.ThrowExceptionIfReadOnly();
            if (this._relationships != null)
                this._relationships.RemoveRelationship(id);
        }

        /// <summary>
        /// Retrieve all the relationships attached to this part.
        /// </summary>
        /// <returns>This part's relationships.</returns>
        /// <exception cref="OpenXml4NetException"></exception>
        /// @see OPC.RelationshipSource#getRelationships()
        public PackageRelationshipCollection Relationships
        {
            get
            {
                return GetRelationshipsCore(null);
            }
        }

        /// <summary>
        /// Retrieves a package relationship from its id.
        /// </summary>
        /// <param name="id">
        /// ID of the package relationship to retrieve.
        /// </param>
        /// <returns>The package relationship</returns>
        /// @see OPC.RelationshipSource#getRelationship(java.lang.String)
        public PackageRelationship GetRelationship(String id)
        {
            return this._relationships.GetRelationshipByID(id);
        }

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
        /// @see OPC.RelationshipSource#getRelationshipsByType(java.lang.String)
        public PackageRelationshipCollection GetRelationshipsByType(
                String relationshipType)
        {
            _container.ThrowExceptionIfWriteOnly();

            return GetRelationshipsCore(relationshipType);
        }

        /// <summary>
        /// Implementation of the getRelationships method().
        /// </summary>
        /// <param name="filter">
        /// Relationship type filter. If <i>null</i> then the filter is
        /// disabled and return all the relationships.
        /// </param>
        /// <returns>All relationships from this part that have the specified type.</returns>
        /// <exception cref="InvalidFormatException">
        /// Throws if an error occurs during parsing the relationships
        /// part.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Throws if the package is open en write only mode.
        /// </exception>
        /// @see #getRelationshipsByType(String)
        private PackageRelationshipCollection GetRelationshipsCore(String filter)
        {
            this._container.ThrowExceptionIfWriteOnly();
            if (_relationships == null)
            {
                this.ThrowExceptionIfRelationship();
                _relationships = new PackageRelationshipCollection(this);
            }
            return new PackageRelationshipCollection(_relationships, filter);
        }

        /// <summary>
        /// Knows if the part have any relationships.
        /// </summary>
        /// <returns><b>true</b> if the part have at least one relationship else
        /// <b>false</b>.
        /// </returns>
        /// @see OPC.RelationshipSource#hasRelationships()
        public bool HasRelationships
        {
            get
            {
                return (!this.IsRelationshipPart && (_relationships != null && _relationships.Size > 0));
            }
        }

        /// <summary>
        /// Checks if the specified relationship is part of this package part.
        /// </summary>
        /// <param name="rel">
        /// The relationship to check.
        /// </param>
        /// <returns><b>true</b> if the specified relationship exists in this part,
        /// else returns <b>false</b>
        /// </returns>
        /// @see OPC.RelationshipSource#isRelationshipExists(OPC.PackageRelationship)
        public bool IsRelationshipExists(PackageRelationship rel)
        {
            foreach (PackageRelationship r in _relationships)
            {
                if (r == rel)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Get the PackagePart that is the target of a relationship.
        /// </summary>
        /// <param name="rel">A relationship from this part to another one</param>
        /// <returns>The target part of the relationship</returns>
        public PackagePart GetRelatedPart(PackageRelationship rel)
        {
            // Ensure this is one of ours
            if (!IsRelationshipExists(rel))
            {
                throw new ArgumentException("Relationship " + rel + " doesn't start with this part " + _partName);
            }

            // Get the target URI, excluding any relative fragments
            Uri target = rel.TargetUri;
            if (target.OriginalString.IndexOf('#') >=0)
            {
                String t = target.ToString();
                try
                {
                    target = PackagingUriHelper.ParseUri(t.Substring(0, t.IndexOf('#')), UriKind.Absolute);
                }
                catch (UriFormatException)
                {
                    throw new InvalidFormatException("Invalid target URI: " + t);
                }
            }

            // Turn that into a name, and fetch
            PackagePartName relName = PackagingUriHelper.CreatePartName(target);
            PackagePart part = _container.GetPart(relName);
            if (part == null)
            {
                throw new ArgumentException("No part found for relationship " + rel);
            }
            return part;
        }


        public Stream GetStream(FileMode mode)
        {
            return this.GetStream(mode, FileAccess.Write);
        }
        public Stream GetStream(FileMode mode, FileAccess access)
        {
            if (mode == FileMode.Create && access == FileAccess.Write)
            {
                return this.GetOutputStream();
            }
            return this.GetInputStream();
        }
        /// <summary>
        /// Get the input stream of this part to read its content.
        /// </summary>
        /// <returns>The input stream of the content of this part, else
        /// <c>null</c>.
        /// </returns>
        public Stream GetInputStream()
        {
            Stream inStream = this.GetInputStreamImpl();
            if (inStream == null)
            {
                throw new IOException("Can't obtain the input stream from "
                        + _partName.Name);
            }
            else
                return inStream;
        }

        /// <summary>
        /// <para>
        /// Get the output stream of this part. If the part is originally embedded in
        /// Zip package, it'll be transform intot a <i>MemoryPackagePart</i> in
        /// order to write inside (the standard Java API doesn't allow to write in
        /// the file)
        /// </para>
        /// <para>
        /// <see cref="MemoryPackagePart" />
        /// </para>
        /// </summary>
        public Stream GetOutputStream()
        {
            Stream outStream;
            // If this part is a zip package part (read only by design) we convert
            // this part into a MemoryPackagePart instance for write purpose.
            if (this is ZipPackagePart)
            {
                // Delete logically this part
                _container.RemovePart(this._partName);

                // Create a memory part
                PackagePart part = _container.CreatePart(this._partName,
                        this._contentType.ToString(), false);
                if (part == null)
                {
                    throw new InvalidOperationException(
                            "Can't create a temporary part !");
                }

                part._relationships = this._relationships;
                outStream = part.GetOutputStreamImpl();
            }
            else
            {
                outStream = this.GetOutputStreamImpl();
            }
            return outStream;
        }


        /// <summary>
        /// Throws an exception if this package part is a relationship part.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// If this part is a relationship part.
        /// </exception>
        private void ThrowExceptionIfRelationship()
        {
            if (this.IsRelationshipPart)
                throw new InvalidOperationException(
                        "Can do this operation on a relationship part !");
        }

        /// <summary>
        /// Ensure the package relationships collection instance is built.
        /// </summary>
        /// <exception cref="InvalidFormatException">
        /// Throws if
        /// </exception>
        private void LoadRelationships()
        {
            if (this._relationships == null && !this.IsRelationshipPart)
            {
                this.ThrowExceptionIfRelationship();
                _relationships = new PackageRelationshipCollection(this);
            }
        }

        /*
         * Accessors
         */

        /// <summary>
        /// </summary>
        /// <returns>the uri</returns>
        public PackagePartName PartName
        {
            get
            {
                return _partName;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>the contentType</returns>
        public String ContentType
        {
            get
            {
                return _contentType.ToString();
            }
            set
            {
                if (_container == null)
                {
                    _contentType = new ContentType(value);
                }
                else
                {
                    _container.UnregisterPartAndContentType(_partName);
                    _contentType = new ContentType(value);
                    _container.RegisterPartAndContentType(this);
                }
            }
        }
        /// <summary>
        /// </summary>
        /// <returns>The Content Type, including parameters, of the part</returns>
        public ContentType ContentTypeDetails
        {
            get
            {
                return _contentType;
            }
        }
        public OPCPackage Package
        {
            get
            {
                return _container;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if this part is a relationship</returns>
        public bool IsRelationshipPart
        {
            get
            {
                return this._isRelationshipPart;
            }
        }

        /// <summary>
        /// </summary>
        /// <returns>true if this part has been logically deleted</returns>
        public bool IsDeleted
        {
            get
            {
                return _isDeleted;
            }
            set { this._isDeleted = value; }
        }
        /// <summary>
        /// </summary>
        /// <returns>The length of the part in bytes, or -1 if not known</returns>
        public virtual long Size
        {
            get
            {
                return -1;
            }
        }
        public override String ToString()
        {
            return "Name: " + this._partName + " - Content Type: "
                    + this._contentType.ToString();
        }

        /// <summary>
        /// Compare based on the package part name, using a natural sort order
        /// </summary>
        public int CompareTo(PackagePart other)
        {
            // NOTE could also throw a NullPointerException() if desired
            if (other == null)
                return -1;

            return PackagePartName.Compare(this._partName, other._partName);
        }
        /*-------------- Abstract methods ------------- */

        /// <summary>
        /// Abtract method that get the input stream of this part.
        /// </summary>
        /// <exception cref="IOException">
        /// Throws if an IO Exception occur in the implementation
        /// method.
        /// </exception>
        protected abstract Stream GetInputStreamImpl();
        /// <summary>
        /// Abstract method that get the output stream of this part.
        /// </summary>
        protected abstract Stream GetOutputStreamImpl();

        /// <summary>
        /// Save the content of this part and the associated relationships part (if
        /// this part own at least one relationship) into the specified output
        /// stream.
        /// </summary>
        /// <param name="zos">
        /// Output stream to save this part.
        /// </param>
        /// <exception cref="OpenXml4NetException">
        /// If any exception occur.
        /// </exception>
        public abstract bool Save(Stream zos);

        /// <summary>
        /// Load the content of this part.
        /// </summary>
        /// <param name="ios">
        /// The input stream of the content to load.
        /// </param>
        /// <returns><b>true</b> if the content has been successfully loaded, else
        /// <b>false</b>.
        /// </returns>
        /// <exception cref="InvalidFormatException">
        /// Throws if the content format is invalid.
        /// </exception>
        public abstract bool Load(Stream ios);

        /// <summary>
        /// Close this part : flush this part, close the input stream and output
        /// stream. After this method call, the part must be available for packaging.
        /// </summary>
        public abstract void Close();

        /// <summary>
        /// Flush the content of this part. If the input stream and/or output stream
        /// as in a waiting state to read or write, the must to empty their
        /// respective buffer.
        /// </summary>
        public abstract void Flush();

        /// <summary>
        /// Allows sub-classes to clean up before new data is added.
        /// </summary>
        public virtual void Clear() { }
    }

}