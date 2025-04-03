using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; 
using System.Xml;
using System.Xml.XPath;

using NPOI.OpenXml4Net.Exceptions;
using NPOI.Util;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Represents a collection of PackageRelationship elements that are owned by a
     * given PackagePart or the Package.
     * 
     * @author Julien Chable, CDubettier
     * @version 0.1
     */
    public class PackageRelationshipCollection : IEnumerator<PackageRelationship>
    {

        private static readonly POILogger logger = POILogFactory.GetLogger(typeof(PackageRelationshipCollection));

        /**
         * Package relationships ordered by ID.
         */
        private readonly SortedList<String, PackageRelationship> relationshipsByID;

        /**
         * A lookup of internal relationships to avoid
         */
        private readonly SortedList<String, PackageRelationship> internalRelationshipsByTargetName;

        /**
         * This relationshipPart.
         */
        private readonly PackagePart relationshipPart;

        /**
         * Source part.
         */
        private readonly PackagePart sourcePart;

        /**
         * This part name.
         */
        private readonly PackagePartName partName;

        /**
         * Reference to the package.
         */
        private readonly OPCPackage container;
        /**
         * The ID number of the next rID# to generate, or -1
         *  if that is still to be determined.
         */
        private int nextRelationshipId = -1;
        /**
         * Constructor.
         */
        public PackageRelationshipCollection()
        {
            relationshipsByID = new SortedList<String, PackageRelationship>();
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
        /**
         * Copy constructor.
         * 
         * This collection will contain only elements from the specified collection
         * for which the type is compatible with the specified relationship type
         * filter.
         * 
         * @param coll
         *            Collection to import.
         * @param filter
         *            Relationship type filter.
         */
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

        /**
         * Constructor.
         */
        public PackageRelationshipCollection(OPCPackage container)
            : this(container, null)
        {

        }

        /**
         * Constructor.
         * 
         * @throws InvalidFormatException
         *             Throws if the format of the content part is invalid.
         * 
         * @throws InvalidOperationException
         *             Throws if the specified part is a relationship part.
         */
        public PackageRelationshipCollection(PackagePart part) :
            this(part._container, part)
        {

        }

        /**
         * Constructor. Parse the existing package relationship part if one exists.
         * 
         * @param container
         *            The parent package.
         * @param part
         *            The part that own this relationships collection. If <b>null</b>
         *            then this part is considered as the package root.
         * @throws InvalidFormatException
         *             If an error occurs during the parsing of the relatinships
         *             part fo the specified part.
         */
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

        /**
         * Get the relationship part name of the specified part.
         * 
         * @param part
         *            The part .
         * @return The relationship part name of the specified part. Be careful,
         *         only the correct name is returned, this method does not check if
         *         the part really exist in a package !
         * @throws InvalidOperationException
         *             Throws if the specified part is a relationship part.
         */
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

        /**
         * Add the specified relationship to the collection.
         * 
         * @param relPart
         *            The relationship to add.
         */
        public void AddRelationship(PackageRelationship relPart)
        {
            if (relPart == null || string.IsNullOrEmpty(relPart.Id))
            {
                throw new ArgumentException("invalid relationship part/id");
            }
            relationshipsByID[relPart.Id] = relPart;
        }

        /**
         * Add a relationship to the collection.
         * 
         * @param targetUri
         *            Target URI.
         * @param targetMode
         *            The target mode : INTERNAL or EXTERNAL
         * @param relationshipType
         *            Relationship type.
         * @param id
         *            Relationship ID.
         * @return The newly created relationship.
         * @see PackageAccess
         */
        public PackageRelationship AddRelationship(Uri targetUri,
                TargetMode targetMode, String relationshipType, String id)
        {
            if (string.IsNullOrEmpty(id))
            {
                // Generate a unique ID if id parameter is null.
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
            if (targetMode == TargetMode.Internal
                && !internalRelationshipsByTargetName.ContainsKey(targetUri.OriginalString))
            {
                internalRelationshipsByTargetName.Add(targetUri.OriginalString, rel);
            }
            return rel;
        }

        /**
         * Remove a relationship by its ID.
         * 
         * @param id
         *            The relationship ID to Remove.
         */
        public void RemoveRelationship(String id)
        {
            if(relationshipsByID == null)
            {
                return;
            }
            PackageRelationship rel = relationshipsByID[id];
            if (rel != null)
            {
                relationshipsByID.Remove(rel.Id);                    
                internalRelationshipsByTargetName.RemoveAt(internalRelationshipsByTargetName.IndexOfValue(rel));
            }
        }

        /**
         * Remove a relationship by its reference.
         * 
         * @param rel
         *            The relationship to delete.
         */
        public void RemoveRelationship(PackageRelationship rel)
        {
            if (rel == null)
                throw new ArgumentException("rel");

            relationshipsByID.Values.Remove(rel);
        }

        /**
         * Retrieves a relationship by its index in the collection.
         * 
         * @param index
         *            Must be a value between [0-relationships_count-1]
         */
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

        /**
         * Retrieves a package relationship based on its id.
         * 
         * @param id
         *            ID of the package relationship to retrieve.
         * @return The package relationship identified by the specified id.
         */
        public PackageRelationship GetRelationshipByID(String id)
        {
            if(id==null)
            {
                throw new ArgumentException("Cannot read relationship, provided ID is empty: " + id +
                    ", having relationships: " + relationshipsByID.Keys.Select(key=>string.Join(",", key)));
            }
            if (!relationshipsByID.TryGetValue(id, out PackageRelationship byId))
                return null;
            return byId;
        }

        /**
         * Get the numbe rof relationships in the collection.
         */
        public int Size
        {
            get
            {
                return relationshipsByID.Values.Count;
            }
        }

        /**
         * Parse the relationship part and add all relationship in this collection.
         * 
         * @param relPart
         *            The package part to parse.
         * @throws InvalidFormatException
         *             Throws if the relationship part is invalid.
         */
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

        /**
         * Retrieves all relations with the specified type.
         * 
         * @param typeFilter
         *            Relationship type filter. If <b>null</b> then all
         *            relationships are returned.
         * @return All relationships of the type specified by the filter.
         */
        public PackageRelationshipCollection GetRelationships(String typeFilter)
        {
            return new PackageRelationshipCollection(this, typeFilter);
        }

        /**
         * Get this collection's iterator.
         */
        public IEnumerator<PackageRelationship> GetEnumerator()
        {
            return relationshipsByID.Values.GetEnumerator();
        }

        /**
         * Get an iterator of a collection with all relationship with the specified
         * type.
         * 
         * @param typeFilter
         *            Type filter.
         * @return An iterator to a collection containing all relationships with the
         *         specified type contain in this collection.
         */
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

        /**
         * Clear all relationships.
         */
        public void Clear()
        {
            relationshipsByID.Clear();
            internalRelationshipsByTargetName.Clear();
        }
        public PackageRelationship FindExistingInternalRelation(PackagePart packagePart)
        {
            var pn=packagePart.PartName.Name;
            if (!internalRelationshipsByTargetName.TryGetValue(pn, out PackageRelationship relation))
                return null;
            return relation;
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
