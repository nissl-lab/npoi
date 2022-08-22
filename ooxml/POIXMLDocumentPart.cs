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
namespace NPOI
{
    using NPOI.Util;
    using NPOI.OpenXml4Net.OPC;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using NPOI.OpenXml4Net.Exceptions;
using System.Xml;
    using NPOI.OpenXml4Net.OPC.Internal;
    using System.Diagnostics;

    /**
     * Represents an entry of a OOXML namespace.
     *
     * <p>
     * Each POIXMLDocumentPart keeps a reference to the underlying a {@link org.apache.poi.openxml4j.opc.PackagePart}.
     * </p>
     *
     * @author Yegor Kozlov
     */
    public class POIXMLDocumentPart
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(POIXMLDocumentPart));
        private String coreDocumentRel = PackageRelationshipTypes.CORE_DOCUMENT;

        private PackagePart packagePart;
        private PackageRelationship packageRel;
        private POIXMLDocumentPart parent;
        private Dictionary<String, RelationPart> relations = new Dictionary<String, RelationPart>();

        /**
         * The RelationPart is a cached relationship between the document, which contains the RelationPart,
         * and one of its referenced child document parts.
         * The child document parts may only belong to one parent, but it's often referenced by other
         * parents too, having varying {@link PackageRelationship#getId() relationship ids} pointing to it.
         */
        public class RelationPart
        {
            private PackageRelationship relationship;
            private POIXMLDocumentPart documentPart;
        
            internal RelationPart(PackageRelationship relationship, POIXMLDocumentPart documentPart)
            {
                this.relationship = relationship;
                this.documentPart = documentPart;
            }

            /**
             * @return the cached relationship, which uniquely identifies this child document part within the parent 
             */
            public PackageRelationship Relationship
            {
                get
                {
                    return relationship;
                }
            }

            /**
             * @return the child document part
             */
            public T GetDocumentPart<T>() where T: POIXMLDocumentPart
            {
                return (T)documentPart;
            }

            public POIXMLDocumentPart DocumentPart
            {
                get
                {
                    return documentPart;
                }
            }
        }

        /**
         * Counter that provides the amount of incoming relations from other parts
         * to this part.
         */
        private int relationCounter = 0;

        int IncrementRelationCounter()
        {
            relationCounter++;
            return relationCounter;
        }

        int DecrementRelationCounter()
        {
            relationCounter--;
            return relationCounter;
        }

        int GetRelationCounter()
        {
            return relationCounter;
        }

        /**
         * Construct POIXMLDocumentPart representing a "core document" namespace part.
         */
        public POIXMLDocumentPart(OPCPackage pkg)
            : this(pkg, PackageRelationshipTypes.CORE_DOCUMENT)
        {
        }

        /**
         * Construct POIXMLDocumentPart representing a custom "core document" package part.
         */
        public POIXMLDocumentPart(OPCPackage pkg, String coreDocumentRel)
            : this(GetPartFromOPCPackage(pkg, coreDocumentRel))
        {
            this.coreDocumentRel = coreDocumentRel;

            PackageRelationship coreRel = pkg.GetRelationshipsByType(this.coreDocumentRel).GetRelationship(0);
            if (coreRel == null)
            {
                coreRel = pkg.GetRelationshipsByType(PackageRelationshipTypes.STRICT_CORE_DOCUMENT).GetRelationship(0);
                if (coreRel != null)
                {
                    throw new POIXMLException("Strict OOXML isn't currently supported, please see bug #57699");
                }
            }
            if (coreRel == null)
            {
                throw new POIXMLException("OOXML file structure broken/invalid - no core document found!");
            }
            this.packagePart = pkg.GetPart(coreRel);
            this.packageRel = coreRel;
        }

        /**
         * Creates new POIXMLDocumentPart   - called by client code to create new parts from scratch.
         *
         * @see #CreateRelationship(POIXMLRelation, POIXMLFactory, int, bool)
         */
        public POIXMLDocumentPart()
        {
        }
        /**
         * Creates an POIXMLDocumentPart representing the given package part and relationship.
         * Called by {@link #read(POIXMLFactory, java.util.Map)} when reading in an existing file.
         *
         * @param part - The package part that holds xml data representing this sheet.
         * @see #read(POIXMLFactory, java.util.Map)
         *
         * @since POI 3.14-Beta1
         */
        public POIXMLDocumentPart(PackagePart part)
            : this(null, part)
        {
            
        }
        /**
         * Creates an POIXMLDocumentPart representing the given package part, relationship and parent
         * Called by {@link #read(POIXMLFactory, java.util.Map)} when reading in an existing file.
         *
         * @param parent - Parent part
         * @param part - The package part that holds xml data representing this sheet.
         * @see #read(POIXMLFactory, java.util.Map)
         *
         * @since POI 3.14-Beta1
         */
        public POIXMLDocumentPart(POIXMLDocumentPart parent, PackagePart part)
        {
            this.packagePart = part;
            this.parent = parent;
        }
        /**
         * Creates an POIXMLDocumentPart representing the given namespace part and relationship.
         * Called by {@link #read(POIXMLFactory, java.util.Map)} when Reading in an exisiting file.
         *
         * @param part - The namespace part that holds xml data represenring this sheet.
         * @param rel - the relationship of the given namespace part
         * @see #read(POIXMLFactory, java.util.Map) 
         */
         [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public POIXMLDocumentPart(PackagePart part, PackageRelationship rel)
            : this(null, part)
        {
        }

        /**
         * Creates an POIXMLDocumentPart representing the given namespace part, relationship and parent
         * Called by {@link #read(POIXMLFactory, java.util.Map)} when Reading in an exisiting file.
         *
         * @param parent - Parent part
         * @param part - The namespace part that holds xml data represenring this sheet.
         * @param rel - the relationship of the given namespace part
         * @see #read(POIXMLFactory, java.util.Map)
         */
         [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public POIXMLDocumentPart(POIXMLDocumentPart parent, PackagePart part, PackageRelationship rel)
            : this(null, part)
        {
            this.packagePart = part;
            this.packageRel = rel;
            this.parent = parent;
        }

        /**
         * When you open something like a theme, call this to
         *  re-base the XML Document onto the core child of the
         *  current core document 
         */
        protected void Rebase(OPCPackage pkg)
        {
            PackageRelationshipCollection cores =
                packagePart.GetRelationshipsByType(coreDocumentRel);
            if (cores.Size != 1)
            {
                throw new InvalidOperationException(
                    "Tried to rebase using " + coreDocumentRel +
                    " but found " + cores.Size + " parts of the right type"
                );
            }
            packagePart = packagePart.GetRelatedPart(cores.GetRelationship(0));
        }
        static XmlNamespaceManager nsm = null;
        public static XmlNamespaceManager NamespaceManager
        {
            get {
                if (nsm == null)
                    nsm = CreateDefaultNSM();
                return nsm;
            }
        }
        internal static XmlNamespaceManager CreateDefaultNSM()
        {
            //  Create a NamespaceManager to handle the default namespace, 
            //  and create a prefix for the default namespace:
            NameTable nt = new NameTable();
            XmlNamespaceManager ns = new XmlNamespaceManager(nt);
            ns.AddNamespace(string.Empty, PackageNamespaces.SCHEMA_MAIN);
            ns.AddNamespace("d", PackageNamespaces.SCHEMA_MAIN);
            ns.AddNamespace("a", PackageNamespaces.SCHEMA_DRAWING);
            ns.AddNamespace("xdr", PackageNamespaces.SCHEMA_SHEETDRAWINGS);
            ns.AddNamespace("r", PackageNamespaces.SCHEMA_RELATIONSHIPS);
            ns.AddNamespace("c", PackageNamespaces.SCHEMA_CHART);
            ns.AddNamespace("vt", PackageNamespaces.SCHEMA_VT);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"); 
            ns.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            ns.AddNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            ns.AddNamespace("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            ns.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            ns.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
            ns.AddNamespace("v", "urn:schemas-microsoft-com:vml");
            ns.AddNamespace("wne","http://schemas.microsoft.com/office/word/2006/wordml");
            // extended properties (app.xml)
            ns.AddNamespace("xp", PackageRelationshipTypes.EXTENDED_PROPERTIES);
            // custom properties
            ns.AddNamespace("ctp", PackageRelationshipTypes.CUSTOM_PROPERTIES);
            // core properties
            ns.AddNamespace("cp", PackagePropertiesPart.NAMESPACE_CP_URI);
            // core property namespaces 
            ns.AddNamespace("dc", PackagePropertiesPart.NAMESPACE_DC_URI);
            ns.AddNamespace("dcterms", PackagePropertiesPart.NAMESPACE_DCTERMS_URI);
            ns.AddNamespace("dcmitype", PackageNamespaces.DCMITYPE);
            ns.AddNamespace("xsi", PackagePropertiesPart.NAMESPACE_XSI_URI);

            ns.AddNamespace("xsd", "http://www.w3.org/2001/XMLSchema");
            ns.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            return ns;
        }
        /**
         * Provides access to the underlying PackagePart
         *
         * @return the underlying PackagePart
         */
        public PackagePart GetPackagePart()
        {
            return packagePart;
        }

        public static XmlDocument ConvertStreamToXml(Stream xmlStream)
        {
            XmlDocument xmlDoc = new XmlDocument();
            NPOI.OpenXml4Net.Util.XmlHelper.LoadXmlSafe(xmlDoc, xmlStream);
            return xmlDoc;
        }

        /**
         * Provides access to the PackageRelationship that identifies this POIXMLDocumentPart
         *
         * @return the PackageRelationship that identifies this POIXMLDocumentPart
         */
         [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public PackageRelationship GetPackageRelationship()
        {
            if (this.parent != null)
            {
                foreach (RelationPart rp in parent.RelationParts)
                {
                    if (rp.DocumentPart == this)
                    {
                        return rp.Relationship;
                    }
                }
            }
            else
            {
                OPCPackage pkg = GetPackagePart().Package;
                String partName = GetPackagePart().PartName.Name;
                foreach (PackageRelationship rel in pkg.Relationships)
                {
                    if (rel.TargetUri.OriginalString.Equals(partName))
                    {
                        return rel;
                    }
                }
            }
            return null;
        }

        /**
         * Returns the list of child relations for this POIXMLDocumentPart
         *
         * @return child relations
         */
        public List<POIXMLDocumentPart> GetRelations()
        {
            List<POIXMLDocumentPart> l = new List<POIXMLDocumentPart>();
            foreach (RelationPart rp in relations.Values)
            {
                l.Add(rp.DocumentPart);
            }
            return l;
        }

        /**
         * Returns the list of child relations for this POIXMLDocumentPart
         *
         * @return child relations
         */
        public List<RelationPart> RelationParts
        {
            get
            {
                List<RelationPart> l = new List<RelationPart>(relations.Values);
                return l;
            }
        }

        /**
         * Returns the target {@link POIXMLDocumentPart}, where a
         * {@link PackageRelationship} is set from the {@link PackagePart} of this
         * {@link POIXMLDocumentPart} to the {@link PackagePart} of the target
         * {@link POIXMLDocumentPart} with a {@link PackageRelationship#GetId()}
         * matching the given parameter value.
         * 
         * @param id
         *            The relation id to look for
         * @return the target part of the relation, or null, if none exists
         */
        public POIXMLDocumentPart GetRelationById(String id)
        {
            if (string.IsNullOrEmpty(id)|| !relations.ContainsKey(id))
                return null;

            RelationPart rp = relations[id];
            return (rp == null) ? null : rp.DocumentPart;
            
        }

        /**
         * Returns the {@link PackageRelationship#GetId()} of the
         * {@link PackageRelationship}, that sources from the {@link PackagePart} of
         * this {@link POIXMLDocumentPart} to the {@link PackagePart} of the given
         * parameter value.
         * 
         * @param part
         *            The {@link POIXMLDocumentPart} for which the according
         *            relation-id shall be found.
         * @return The value of the {@link PackageRelationship#GetId()} or null, if
         *         parts are not related.
         */
        public String GetRelationId(POIXMLDocumentPart part)
        {
            foreach (KeyValuePair<String, RelationPart> entry in relations)
            {
                if (entry.Value.DocumentPart == part)
                {
                    return entry.Value.Relationship.Id;
                }
            }
            return null;
        }

        /// <summary>
        /// Add a new child POIXMLDocumentPart
        /// </summary>
        /// <param name="id"></param>
        /// <param name="part">the child to add</param>
        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public void AddRelation(String id, POIXMLDocumentPart part)
        {
            PackageRelationship pr = part.GetPackagePart().GetRelationship(id);
            AddRelation(pr, part);
        }
        /// <summary>
        /// Add a new child POIXMLDocumentPart
        /// </summary>
        /// <param name="relId">the preferred relation id, when null the next free relation id will be used</param>
        /// <param name="relationshipType">the package relationship type</param>
        /// <param name="part">the child to add</param>
        /// <returns></returns>
        public RelationPart AddRelation(String relId, POIXMLRelation relationshipType, POIXMLDocumentPart part)
        {
            PackageRelationship pr = this.packagePart.FindExistingRelation(part.GetPackagePart()); 
            if (pr == null)
            {
                PackagePartName ppn = part.GetPackagePart().PartName;
                String relType = relationshipType.Relation;
                pr = packagePart.AddRelationship(ppn, TargetMode.Internal, relType, relId);
            }
            AddRelation(pr, part);
            return new RelationPart(pr, part);
        }
        /// <summary>
        /// Add a new child POIXMLDocumentPart
        /// </summary>
        /// <param name="pr">the relationship of the child</param>
        /// <param name="part">the child to add</param>
        private void AddRelation(PackageRelationship pr, POIXMLDocumentPart part)
        {
            if (relations.ContainsKey(pr.Id))
                relations[pr.Id] = new RelationPart(pr, part);
            else
                relations.Add(pr.Id, new RelationPart(pr, part));
            part.IncrementRelationCounter();

        }

        /**
         * Remove the relation to the specified part in this namespace and remove the
         * part, if it is no longer needed.
         */
        protected internal void RemoveRelation(POIXMLDocumentPart part)
        {
            RemoveRelation(part, true);
        }

        /**
         * Remove the relation to the specified part in this namespace and remove the
         * part, if it is no longer needed and flag is set to true.
         * 
         * @param part
         *            The related part, to which the relation shall be Removed.
         * @param RemoveUnusedParts
         *            true, if the part shall be Removed from the namespace if not
         *            needed any longer.
         */
        protected internal bool RemoveRelation(POIXMLDocumentPart part, bool RemoveUnusedParts)
        {
            String id = GetRelationId(part);
            if (id == null)
            {
                // part is not related with this POIXMLDocumentPart
                return false;
            }
            /* decrement usage counter */
            part.DecrementRelationCounter();
            /* remove namespacepart relationship */
            GetPackagePart().RemoveRelationship(id);
            /* remove POIXMLDocument from relations */
            relations.Remove(id);

            if (RemoveUnusedParts)
            {
                /* if last relation to target part was Removed, delete according target part */
                if (part.GetRelationCounter() == 0)
                {
                    try
                    {
                        part.onDocumentRemove();
                    }
                    catch (IOException e)
                    {
                        throw new POIXMLException(e);
                    }
                    GetPackagePart().Package.RemovePart(part.GetPackagePart());
                }
            }
            return true;
        }

        /**
         * Returns the parent POIXMLDocumentPart. All parts except root have not-null parent.
         *
         * @return the parent POIXMLDocumentPart or <code>null</code> for the root element.
         */
        public POIXMLDocumentPart GetParent()
        {
            return parent;
        }

        public override String ToString()
        {
            return packagePart == null ? string.Empty : packagePart.ToString();
        }

        /**
         * Save the content in the underlying namespace part.
         * Default implementation is empty meaning that the namespace part is left unmodified.
         *
         * Sub-classes should override and add logic to marshal the "model" into Ooxml4J.
         *
         * For example, the code saving a generic XML entry may look as follows:
         * <pre><code>
         * protected void commit()  {
         *   PackagePart part = GetPackagePart();
         *   Stream out = part.GetStream();
         *   XmlObject bean = GetXmlBean(); //the "model" which holds Changes in memory
         *   bean.save(out, DEFAULT_XML_OPTIONS);
         *   out.close();
         * }
         *  </code></pre>
         *
         */
        protected internal virtual void Commit()
        {

        }

        /**
         * Save Changes in the underlying OOXML namespace.
         * Recursively fires {@link #commit()} for each namespace part
         *
         * @param alreadySaved    context set Containing already visited nodes
         */
        protected internal void OnSave(List<PackagePart> alreadySaved)
        {
            // this usually clears out previous content in the part...
            PrepareForCommit();

            Commit();
            alreadySaved.Add(this.GetPackagePart());
            foreach (RelationPart rp in relations.Values)
            {
                POIXMLDocumentPart p = rp.DocumentPart;
                if (!alreadySaved.Contains(p.GetPackagePart()))
                {
                    p.OnSave(alreadySaved);
                }
            }
        }
        /**
         * Ensure that a memory based package part does not have lingering data from previous 
         * commit() calls. 
         * 
         * Note: This is overwritten for some objects, as *PictureData seem to store the actual content 
         * in the part directly without keeping a copy like all others therefore we need to handle them differently.
         */
        protected internal virtual void PrepareForCommit()
        {
            PackagePart part = this.GetPackagePart();
            if (part != null)
            {
                part.Clear();
            }
        }
        /**
         * Create a new child POIXMLDocumentPart
         *
         * @param descriptor the part descriptor
         * @param factory the factory that will create an instance of the requested relation
         * @return the Created child POIXMLDocumentPart
         */
        public POIXMLDocumentPart CreateRelationship(POIXMLRelation descriptor, POIXMLFactory factory)
        {
            return CreateRelationship(descriptor, factory, -1, false).DocumentPart;
        }

        public POIXMLDocumentPart CreateRelationship(POIXMLRelation descriptor, POIXMLFactory factory, int idx)
        {
            return CreateRelationship(descriptor, factory, idx, false).DocumentPart;
        }

        /**
         * Create a new child POIXMLDocumentPart
         *
         * @param descriptor the part descriptor
         * @param factory the factory that will create an instance of the requested relation
         * @param idx part number
         * @param noRelation if true, then no relationship is Added.
         * @return the Created child POIXMLDocumentPart
         */
        protected RelationPart CreateRelationship(POIXMLRelation descriptor, POIXMLFactory factory, int idx, bool noRelation)
        {
            try
            {
                PackagePartName ppName = PackagingUriHelper.CreatePartName(descriptor.GetFileName(idx));
                PackageRelationship rel = null;
                PackagePart part = packagePart.Package.CreatePart(ppName, descriptor.ContentType);
                if (!noRelation)
                {
                    /* only add to relations, if according relationship is being Created. */
                    rel = packagePart.AddRelationship(ppName, TargetMode.Internal, descriptor.Relation);
                }
                POIXMLDocumentPart doc = factory.NewDocumentPart(descriptor);
                doc.packageRel = rel;
                doc.packagePart = part;
                doc.parent = this;
                if (!noRelation)
                {
                    /* only add to relations, if according relationship is being Created. */
                    AddRelation(rel, doc);
                }
                return new RelationPart(rel, doc);
            }
            catch (PartAlreadyExistsException pae)
            {
                // Return the specific exception so the user knows
                //  that the name is already taken
                throw pae;
            }
            catch (Exception e)
            {
                // Give a general wrapped exception for the problem
                throw new POIXMLException(e);
            }
        }

        public TValue PutDictionary<TKey, TValue>(Dictionary<TKey, TValue> dict, TKey key, TValue value)
        {
            TValue oldValue = default(TValue);
            if (dict.ContainsKey(key))
            {
                oldValue = dict[key];
                dict[key] = value;
            }
            else
                dict.Add(key, value);
            return oldValue;
        }
        public TValue GetDictionary<TKey, TValue>(Dictionary<TKey, TValue> dict, TKey key)
        {
            if (dict.ContainsKey(key))
            {
                return dict[key];
            }
            return default(TValue);
        }
        /**
         * Iterate through the underlying PackagePart and create child POIXMLFactory instances
         * using the specified factory
         *
         * @param factory   the factory object that Creates POIXMLFactory instances
         * @param context   context map Containing already visited noted keyed by tarGetURI
         */
        protected void Read(POIXMLFactory factory, Dictionary<PackagePart, POIXMLDocumentPart> context)
        {
            PackagePart pp = GetPackagePart();
            // add mapping a second time, in case of initial caller hasn't done so
            POIXMLDocumentPart otherChild = PutDictionary(context, pp, this);
            if (otherChild != null && otherChild != this)
            {
                throw new POIXMLException("Unique PackagePart-POIXMLDocumentPart relation broken!");
            }

            if (!pp.HasRelationships) return;

            PackageRelationshipCollection rels = packagePart.Relationships;
            List<POIXMLDocumentPart> readLater = new List<POIXMLDocumentPart>();

            // scan breadth-first, so parent-relations are hopefully the shallowest element
            foreach (PackageRelationship rel in rels)
            {
                if (rel.TargetMode == TargetMode.Internal)
                {
                    Uri uri = rel.TargetUri;

                    // check for internal references (e.g. '#Sheet1!A1')
                    PackagePartName relName;
                    //if (uri.getRawFragment() != null)
                    if (uri.OriginalString.IndexOf('#') >= 0)
                    {
                        string path = string.Empty;
                        try
                        {
                            path = uri.AbsolutePath;
                        }
                        catch (InvalidOperationException)
                        {
                            path = uri.OriginalString.Substring(0, uri.OriginalString.IndexOf('#'));
                        }
                        relName = PackagingUriHelper.CreatePartName(path);
                    }
                    else
                    {
                        relName = PackagingUriHelper.CreatePartName(uri);
                    }

                    PackagePart p = packagePart.Package.GetPart(relName);
                    if (p == null)
                    {
                        //logger.log(POILogger.ERROR, "Skipped invalid entry " + rel.TargetUri);
                        continue;
                    }

                    POIXMLDocumentPart childPart = GetDictionary(context, p);
                    if (childPart == null)
                    {
                        childPart = factory.CreateDocumentPart(this, p);
                        childPart.parent = this;
                        // already add child to context, so other children can reference it
                        PutDictionary(context, p, childPart);
                        readLater.Add(childPart);
                    }

                    AddRelation(rel, childPart);
                }
            }

            foreach (POIXMLDocumentPart childPart in readLater)
            {
                childPart.Read(factory, context);
            }
        }
        /**
         * Get the PackagePart that is the target of a relationship.
         *
         * @param rel The relationship
         * @return The target part
         * @throws InvalidFormatException
         */
        protected PackagePart GetTargetPart(PackageRelationship rel)
        {
            return GetPackagePart().GetRelatedPart(rel);
        }
        /**
         * Fired when a new namespace part is Created
         */
        internal virtual void OnDocumentCreate()
        {

        }

        /**
         * Fired when a namespace part is read
         */
        internal virtual void OnDocumentRead()
        {

        }



        /**
         * Fired when a namespace part is about to be Removed from the namespace
         */
        protected virtual void onDocumentRemove()
        {

        }

        /**
         * Retrieves the core document part
         * 
         * @since POI 3.14-Beta1
         */
        private static PackagePart GetPartFromOPCPackage(OPCPackage pkg, String coreDocumentRel)
        {
            PackageRelationship coreRel = pkg.GetRelationshipsByType(coreDocumentRel).GetRelationship(0);

            if (coreRel != null)
            {
                PackagePart pp = pkg.GetPart(coreRel);
                if (pp == null)
                {
                    throw new POIXMLException("OOXML file structure broken/invalid - core document '" + coreRel.TargetUri + "' not found.");
                }
                return pp;
            }

            coreRel = pkg.GetRelationshipsByType(PackageRelationshipTypes.STRICT_CORE_DOCUMENT).GetRelationship(0);
            if (coreRel != null)
            {
                throw new POIXMLException("Strict OOXML isn't currently supported, please see bug #57699");
            }

            throw new POIXMLException("OOXML file structure broken/invalid - no core document found!");
        }
    }
}






