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
using System.Xml;
using System.Xml.XPath;

using NPOI.OpenXml4Net.Exceptions;
using NPOI.Util;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXml4Net.OPC
{
    /// <summary>
    /// Represents a collection of PackageRelationship elements that are owned by a
    /// given PackagePart or the Package.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable, CDubettier
    /// @version 0.1
    /// </remarks>

    public class PackageRelationshipCollection : IEnumerator<PackageRelationship>
    {

        private static POILogger logger = POILogFactory.GetLogger(typeof(PackageRelationshipCollection));

        /// <summary>
        /// Package relationships ordered by ID.
        /// </summary>
        private SortedList<String, PackageRelationship> relationshipsByID;

        /// <summary>
        /// Package relationships ordered by type.
        /// </summary>
        private SortedList<String, PackageRelationship> relationshipsByType;
        /// <summary>
        /// A lookup of internal relationships to avoid
        /// </summary>
        private SortedList<String, PackageRelationship> internalRelationshipsByTargetName;

        /// <summary>
        /// This relationshipPart.
        /// </summary>
        private PackagePart relationshipPart;

        /// <summary>
        /// Source part.
        /// </summary>
        private PackagePart sourcePart;

        /// <summary>
        /// This part name.
        /// </summary>
        private PackagePartName partName;

        /// <summary>
        /// Reference to the package.
        /// </summary>
        private OPCPackage container;
        /// <summary>
        /// The ID number of the next rID# to generate, or -1
        ///  if that is still to be determined.
        /// </summary>
        private int nextRelationshipId = -1;
        /// <summary>
        /// Constructor.
        /// </summary>
        public PackageRelationshipCollection()
        {
            relationshipsByID = new SortedList<String, PackageRelationship>();
            relationshipsByType = new SortedList<String, PackageRelationship>(new DuplicateComparer());
            internalRelationshipsByTargetName = new SortedList<string, PackageRelationship>();
        }
        class DuplicateComparer : IComparer<string>
        {

            #region IComparer<string> Members

            public int Compare(string x, string y)
            {
                if (x.CompareTo(y) < 0)
                {
                    return -1;
                }
                return 1;
            }

            #endregion
        }
        /// <summary>
        /// <para>
        /// Copy constructor.
        /// </para>
        /// <para>
        /// This collection will contain only elements from the specified collection
        /// for which the type is compatible with the specified relationship type
        /// filter.
        /// </para>
        /// </summary>
        /// <param name="coll">
        /// Collection to import.
        /// </param>
        /// <param name="filter">
        /// Relationship type filter.
        /// </param>
        public PackageRelationshipCollection(PackageRelationshipCollection coll,
                String filter)
            : this()
        {

            foreach (PackageRelationship rel in coll.relationshipsByID.Values)
            {
                if (filter == null || rel.RelationshipType.Equals(filter))
                    AddRelationship(rel);
            }
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        public PackageRelationshipCollection(OPCPackage container)
            : this(container, null)
        {

        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <exception cref="InvalidFormatException">
        /// Throws if the format of the content part is invalid.
        /// </exception>
        /// 
        /// <exception cref="InvalidOperationException">
        /// Throws if the specified part is a relationship part.
        /// </exception>
        public PackageRelationshipCollection(PackagePart part) :
            this(part._container, part)
        {

        }

        /// <summary>
        /// Constructor. Parse the existing package relationship part if one exists.
        /// </summary>
        /// <param name="container">
        /// The parent package.
        /// </param>
        /// <param name="part">
        /// The part that own this relationships collection. If <b>null</b>
        /// then this part is considered as the package root.
        /// </param>
        /// <exception cref="InvalidFormatException">
        /// If an error occurs during the parsing of the relatinships
        /// part fo the specified part.
        /// </exception>
        public PackageRelationshipCollection(OPCPackage container, PackagePart part)
            : this()
        {


            if (container == null)
                throw new ArgumentException("container needs to be specified");

            // Check if the specified part is not a relationship part
            if (part != null && part.IsRelationshipPart)
                throw new ArgumentException("part");

            this.container = container;
            this.sourcePart = part;
            this.partName = GetRelationshipPartName(part);
            if ((container.GetPackageAccess() != PackageAccess.WRITE)
                    && container.ContainPart(this.partName))
            {
                relationshipPart = container.GetPart(this.partName);
                ParseRelationshipsPart(relationshipPart);
            }
        }

        /// <summary>
        /// Get the relationship part name of the specified part.
        /// </summary>
        /// <param name="part">
        /// The part .
        /// </param>
        /// <returns>The relationship part name of the specified part. Be careful,
        /// only the correct name is returned, this method does not check if
        /// the part really exist in a package !
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Throws if the specified part is a relationship part.
        /// </exception>
        private static PackagePartName GetRelationshipPartName(PackagePart part)
        {
            PackagePartName partName;
            if (part == null)
            {
                partName = PackagingUriHelper.PACKAGE_ROOT_PART_NAME;
            }
            else
            {
                partName = part.PartName;
            }
            return PackagingUriHelper.GetRelationshipPartName(partName);
        }

        /// <summary>
        /// Add the specified relationship to the collection.
        /// </summary>
        /// <param name="relPart">
        /// The relationship to add.
        /// </param>
        public void AddRelationship(PackageRelationship relPart)
        {
            if (relPart == null || string.IsNullOrEmpty(relPart.Id))
            {
                throw new ArgumentException("invalid relationship part/id");
            }
            relationshipsByID[relPart.Id] = relPart;
            relationshipsByType[relPart.RelationshipType] = relPart;
        }

        /// <summary>
        /// Add a relationship to the collection.
        /// </summary>
        /// <param name="targetUri">
        /// Target URI.
        /// </param>
        /// <param name="targetMode">
        /// The target mode : INTERNAL or EXTERNAL
        /// </param>
        /// <param name="relationshipType">
        /// Relationship type.
        /// </param>
        /// <param name="id">
        /// Relationship ID.
        /// </param>
        /// <returns>The newly created relationship.</returns>
        /// @see PackageAccess
        public PackageRelationship AddRelationship(Uri targetUri,
                TargetMode targetMode, String relationshipType, String id)
        {
            if (id == null)
            {
                // Generate a unique ID is id parameter is null.
                if (nextRelationshipId == -1)
                {
                    nextRelationshipId = Size + 1;
                }

                // Work up until we find a unique number (there could be gaps etc)
                do
                {
                    id = "rId" + nextRelationshipId++;
                } while (relationshipsByID.ContainsKey(id));
            }

            PackageRelationship rel = new PackageRelationship(container,
                    sourcePart, targetUri, targetMode, relationshipType, id);
            relationshipsByID[rel.Id] = rel;
            relationshipsByType[rel.RelationshipType] = rel;
            if (targetMode == TargetMode.Internal
                && !internalRelationshipsByTargetName.ContainsKey(targetUri.OriginalString))
            {
                internalRelationshipsByTargetName.Add(targetUri.OriginalString, rel);
            }
            return rel;
        }

        /// <summary>
        /// Remove a relationship by its ID.
        /// </summary>
        /// <param name="id">
        /// The relationship ID to Remove.
        /// </param>
        public void RemoveRelationship(String id)
        {
            if (relationshipsByID != null && relationshipsByType != null)
            {
                PackageRelationship rel = relationshipsByID[id];
                if (rel != null)
                {
                    relationshipsByID.Remove(rel.Id);
                    for (int i = 0; i < relationshipsByType.Count; i++)
                    {
                        if (relationshipsByType.Values[i] == rel)
                            relationshipsByType.RemoveAt(i);
                    }
                    
                    internalRelationshipsByTargetName.RemoveAt(internalRelationshipsByTargetName.IndexOfValue(rel));
                }
            }
        }

        /// <summary>
        /// Remove a relationship by its reference.
        /// </summary>
        /// <param name="rel">
        /// The relationship to delete.
        /// </param>
        public void RemoveRelationship(PackageRelationship rel)
        {
            if (rel == null)
                throw new ArgumentException("rel");

            relationshipsByID.Values.Remove(rel);
            relationshipsByType.Values.Remove(rel);
        }

        /// <summary>
        /// Retrieves a relationship by its index in the collection.
        /// </summary>
        /// <param name="index">
        /// Must be a value between [0-relationships_count-1]
        /// </param>
        public PackageRelationship GetRelationship(int index)
        {
            if (index < 0 || index > relationshipsByID.Values.Count)
                throw new ArgumentException("index");

            int i = 0;
            foreach (PackageRelationship rel in relationshipsByID.Values)
            {
                if (index == i++)
                    return rel;
            }
            return null;
        }

        /// <summary>
        /// Retrieves a package relationship based on its id.
        /// </summary>
        /// <param name="id">
        /// ID of the package relationship to retrieve.
        /// </param>
        /// <returns>The package relationship identified by the specified id.</returns>
        public PackageRelationship GetRelationshipByID(String id)
        {
            if (!relationshipsByID.ContainsKey(id))
                return null;
            return relationshipsByID[id];
        }

        /// <summary>
        /// Get the numbe rof relationships in the collection.
        /// </summary>
        public int Size
        {
            get
            {
                return relationshipsByID.Values.Count;
            }
        }

        /// <summary>
        /// Parse the relationship part and add all relationship in this collection.
        /// </summary>
        /// <param name="relPart">
        /// The package part to parse.
        /// </param>
        /// <exception cref="InvalidFormatException">
        /// Throws if the relationship part is invalid.
        /// </exception>
        private void ParseRelationshipsPart(PackagePart relPart)
        {
            try
            {
                logger.Log(POILogger.DEBUG, "Parsing relationship: " + relPart.PartName);
                XPathDocument xmlRelationshipsDoc = DocumentHelper.ReadDocument(relPart.GetInputStream());

                // Check OPC compliance M4.1 rule
                bool fCorePropertiesRelationship = false;
                //xmlRelationshipsDoc.ChildNodes.GetEnumerator();
                XPathNavigator xpathnav = xmlRelationshipsDoc.CreateNavigator();
                XmlNamespaceManager nsMgr = new XmlNamespaceManager(xpathnav.NameTable);
                nsMgr.AddNamespace("x", PackageNamespaces.RELATIONSHIPS);

                XPathNodeIterator iterator = xpathnav.Select("//x:" + PackageRelationship.RELATIONSHIP_TAG_NAME, nsMgr);

                while (iterator.MoveNext())
                {
                    // Relationship ID
                    String id = iterator.Current.GetAttribute(PackageRelationship.ID_ATTRIBUTE_NAME, xpathnav.NamespaceURI);
                    // Relationship type
                    String type = iterator.Current.GetAttribute(
                            PackageRelationship.TYPE_ATTRIBUTE_NAME, xpathnav.NamespaceURI);

                    /* Check OPC Compliance */
                    // Check Rule M4.1
                    if (type.Equals(PackageRelationshipTypes.CORE_PROPERTIES))
                        if (!fCorePropertiesRelationship)
                            fCorePropertiesRelationship = true;
                        else
                            throw new InvalidFormatException(
                                    "OPC Compliance error [M4.1]: there is more than one core properties relationship in the package !");

                    /* End OPC Compliance */

                    // TargetMode (default value "Internal")
                    string targetModeAttr = iterator.Current.GetAttribute(PackageRelationship.TARGET_MODE_ATTRIBUTE_NAME, xpathnav.NamespaceURI);
                    TargetMode targetMode = TargetMode.Internal;
                    if (targetModeAttr != string.Empty)
                    {
                        targetMode = targetModeAttr.ToLower()
                                .Equals("internal") ? TargetMode.Internal
                                : TargetMode.External;
                    }

                    // Target converted in URI
                    Uri target = PackagingUriHelper.ToUri("http://invalid.uri"); // dummy url
                    String value = iterator.Current.GetAttribute(
                                PackageRelationship.TARGET_ATTRIBUTE_NAME, xpathnav.NamespaceURI); ;
                    try
                    {
                        // when parsing of the given uri fails, we can either
                        // ignore this relationship, which leads to InvalidOperationException
                        // later on, or use a dummy value and thus enable processing of the
                        // package
                        target = PackagingUriHelper.ToUri(value);
                    }
                    catch (UriFormatException e)
                    {
                        logger.Log(POILogger.ERROR, "Cannot convert " + value
                                + " in a valid relationship URI-> dummy-URI used", e);
                    }
                    AddRelationship(target, targetMode, type, id);
                }
            }
            catch (Exception e)
            {
                logger.Log(POILogger.ERROR, e);
                throw new InvalidFormatException(e.Message);
            }
        }

        /// <summary>
        /// Retrieves all relations with the specified type.
        /// </summary>
        /// <param name="typeFilter">
        /// Relationship type filter. If <b>null</b> then all
        /// relationships are returned.
        /// </param>
        /// <returns>All relationships of the type specified by the filter.</returns>
        public PackageRelationshipCollection GetRelationships(String typeFilter)
        {
            return new PackageRelationshipCollection(this, typeFilter);
        }

        /// <summary>
        /// Get this collection's iterator.
        /// </summary>
        public IEnumerator<PackageRelationship> GetEnumerator()
        {
            return relationshipsByID.Values.GetEnumerator();
        }

        /// <summary>
        /// Get an iterator of a collection with all relationship with the specified
        /// type.
        /// </summary>
        /// <param name="typeFilter">
        /// Type filter.
        /// </param>
        /// <returns>An iterator to a collection containing all relationships with the
        /// specified type contain in this collection.
        /// </returns>
        public IEnumerator<PackageRelationship> Iterator(String typeFilter)
        {
            List<PackageRelationship> retArr = new List<PackageRelationship>();
            foreach (PackageRelationship rel in relationshipsByID.Values)
            {
                if (rel.RelationshipType.Equals(typeFilter))
                    retArr.Add(rel);
            }
            return retArr.GetEnumerator();
        }

        /// <summary>
        /// Clear all relationships.
        /// </summary>
        public void Clear()
        {
            relationshipsByID.Clear();
            relationshipsByType.Clear();
            internalRelationshipsByTargetName.Clear();
        }
        public PackageRelationship FindExistingInternalRelation(PackagePart packagePart)
        {
            var pn=packagePart.PartName.Name;
            if (!internalRelationshipsByTargetName.ContainsKey(pn))
                return null;
            return internalRelationshipsByTargetName[pn];
        }
        public override String ToString()
        {
            String str;
            if (relationshipsByID == null)
            {
                str = "relationshipsByID=null";
            }
            else
            {
                str = relationshipsByID.Count + " relationship(s) = [";
            }
            if ((relationshipPart != null) && (relationshipPart.PartName != null))
            {
                str = str + "," + relationshipPart.PartName;
            }
            else
            {
                str = str + ",relationshipPart=null";
            }

            // Source of this relationship
            if ((sourcePart != null) && (sourcePart.PartName != null))
            {
                str = str + "," + sourcePart.PartName;
            }
            else
            {
                str = str + ",sourcePart=null";
            }
            if (partName != null)
            {
                str = str + "," + partName;
            }
            else
            {
                str = str + ",uri=null)";
            }
            return str + "]";
        }

        #region IEnumerator<PackageRelationship> Members

        PackageRelationship IEnumerator<PackageRelationship>.Current
        {
            get { throw new NotImplementedException(); }
        }

        #endregion

        #region IDisposable Members

        void IDisposable.Dispose()
        {
            //relationshipsByID=null;
            //relationshipsByType = null;
        }

        #endregion

        #region IEnumerator Members

        object System.Collections.IEnumerator.Current
        {
            get { throw new NotImplementedException(); }
        }

        bool System.Collections.IEnumerator.MoveNext()
        {
            throw new NotImplementedException();
        }

        void System.Collections.IEnumerator.Reset()
        {
            Clear();
        }

        #endregion
    }

}
