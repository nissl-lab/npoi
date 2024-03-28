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

    /// <summary>
    /// <para>
    /// Represents an entry of a OOXML namespace.
    /// </para>
    /// <para>
    /// 
    /// Each POIXMLDocumentPart keeps a reference to the underlying a <see cref="org.apache.poi.openxml4j.opc.PackagePart" />.
    /// </para>
    /// </summary>
    /// @author Yegor Kozlov

    public class POIXMLDocumentPart
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(POIXMLDocumentPart));

        private String coreDocumentRel = PackageRelationshipTypes.CORE_DOCUMENT;
        private PackagePart packagePart;
        private POIXMLDocumentPart parent;
        private Dictionary<String, RelationPart> relations = new Dictionary<String, RelationPart>();

        /// <summary>
        /// The RelationPart is a cached relationship between the document, which contains the RelationPart,
        /// and one of its referenced child document parts.
        /// The child document parts may only belong to one parent, but it's often referenced by other
        /// parents too, having varying <see cref="PackageRelationship.getId() relationship ids" /> pointing to it.
        /// </summary>

        public class RelationPart
        {
            private PackageRelationship relationship;
            private POIXMLDocumentPart documentPart;

            internal RelationPart(PackageRelationship relationship, POIXMLDocumentPart documentPart)
            {
                this.relationship = relationship;
                this.documentPart = documentPart;
            }

            /// <summary>
            /// </summary>
            /// <returns>the cached relationship, which uniquely identifies this child document part within the parent</returns>

            public PackageRelationship Relationship
            {
                get
                {
                    return relationship;
                }
            }

            /// <summary>
            /// </summary>
            /// <returns>the child document part</returns>

            public T GetDocumentPart<T>() where T : POIXMLDocumentPart
            {
                return (T) documentPart;
            }

            public POIXMLDocumentPart DocumentPart
            {
                get
                {
                    return documentPart;
                }
            }
        }

        /// <summary>
        /// Counter that provides the amount of incoming relations from other parts
        /// to this part.
        /// </summary>
        private int relationCounter = 0;

        private int IncrementRelationCounter()
        {
            relationCounter++;
            return relationCounter;
        }

        private int DecrementRelationCounter()
        {
            relationCounter--;
            return relationCounter;
        }

        private int GetRelationCounter()
        {
            return relationCounter;
        }

        /// <summary>
        /// Construct POIXMLDocumentPart representing a "core document" namespace part.
        /// </summary>

        public POIXMLDocumentPart(OPCPackage pkg)
            : this(pkg, PackageRelationshipTypes.CORE_DOCUMENT)
        {
        }

        /// <summary>
        /// Construct POIXMLDocumentPart representing a custom "core document" package part.
        /// </summary>

        public POIXMLDocumentPart(OPCPackage pkg, String coreDocumentRel)
            : this(GetPartFromOPCPackage(pkg, coreDocumentRel))
        {
            this.coreDocumentRel = coreDocumentRel;
        }

        /// <summary>
        /// Creates new POIXMLDocumentPart   - called by client code to create new parts from scratch.
        /// </summary>
        /// <see cref="CreateRelationship(POIXMLRelation, POIXMLFactory, int, bool)" />

        public POIXMLDocumentPart()
        {
        }

        /// <summary>
        /// Creates an POIXMLDocumentPart representing the given package part and relationship.
        /// Called by <see cref="read(POIXMLFactory, java.util.Map)" /> when reading in an existing file.
        /// </summary>
        /// <param name="part">- The package part that holds xml data representing this sheet.</param>
        /// <see cref="read(POIXMLFactory, java.util.Map)" />
        ///
        /// @since POI 3.14-Beta1

        public POIXMLDocumentPart(PackagePart part)
            : this(null, part)
        {
        }

        /// <summary>
        /// Creates an POIXMLDocumentPart representing the given package part, relationship and parent
        /// Called by <see cref="read(POIXMLFactory, java.util.Map)" /> when reading in an existing file.
        /// </summary>
        /// <param name="parent">- Parent part</param>
        /// <param name="part">- The package part that holds xml data representing this sheet.</param>
        /// <see cref="read(POIXMLFactory, java.util.Map)" />
        ///
        /// @since POI 3.14-Beta1

        public POIXMLDocumentPart(POIXMLDocumentPart parent, PackagePart part)
        {
            this.packagePart = part;
            this.parent = parent;
        }

        /// <summary>
        /// Creates an POIXMLDocumentPart representing the given namespace part and relationship.
        /// Called by <see cref="read(POIXMLFactory, java.util.Map)" /> when Reading in an exisiting file.
        /// </summary>
        /// <param name="part">- The namespace part that holds xml data represenring this sheet.</param>
        /// <param name="rel">- the relationship of the given namespace part</param>
        /// <see cref="read(POIXMLFactory, java.util.Map)" />

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public POIXMLDocumentPart(PackagePart part, PackageRelationship rel)
           : this(null, part)
        {
        }

        /// <summary>
        /// Creates an POIXMLDocumentPart representing the given namespace part, relationship and parent
        /// Called by <see cref="read(POIXMLFactory, java.util.Map)" /> when Reading in an exisiting file.
        /// </summary>
        /// <param name="parent">- Parent part</param>
        /// <param name="part">- The namespace part that holds xml data represenring this sheet.</param>
        /// <param name="rel">- the relationship of the given namespace part</param>
        /// <see cref="read(POIXMLFactory, java.util.Map)" />

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public POIXMLDocumentPart(POIXMLDocumentPart parent, PackagePart part, PackageRelationship rel)
           : this(null, part)
        {
        }

        /// <summary>
        /// When you open something like a theme, call this to
        ///  re-base the XML Document onto the core child of the
        ///  current core document
        /// </summary>

        protected void Rebase(OPCPackage pkg)
        {
            PackageRelationshipCollection cores =
                packagePart.GetRelationshipsByType(coreDocumentRel);
            if(cores.Size != 1)
            {
                throw new InvalidOperationException(
                    "Tried to rebase using " + coreDocumentRel +
                    " but found " + cores.Size + " parts of the right type"
                );
            }
            packagePart = packagePart.GetRelatedPart(cores.GetRelationship(0));
        }

        private static XmlNamespaceManager nsm = null;

        public static XmlNamespaceManager NamespaceManager
        {
            get
            {
                if(nsm == null)
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
            ns.AddNamespace("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
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

        /// <summary>
        /// Provides access to the underlying PackagePart
        /// </summary>
        /// <returns>the underlying PackagePart</returns>

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

        /// <summary>
        /// Provides access to the PackageRelationship that identifies this POIXMLDocumentPart
        /// </summary>
        /// <returns>the PackageRelationship that identifies this POIXMLDocumentPart</returns>

        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public PackageRelationship GetPackageRelationship()
        {
            if(this.parent != null)
            {
                foreach(RelationPart rp in parent.RelationParts)
                {
                    if(rp.DocumentPart == this)
                    {
                        return rp.Relationship;
                    }
                }
            }
            else
            {
                OPCPackage pkg = GetPackagePart().Package;
                String partName = GetPackagePart().PartName.Name;
                foreach(PackageRelationship rel in pkg.Relationships)
                {
                    if(rel.TargetUri.OriginalString.Equals(partName))
                    {
                        return rel;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Returns the list of child relations for this POIXMLDocumentPart
        /// </summary>
        /// <returns>child relations</returns>

        public IList<POIXMLDocumentPart> GetRelations()
        {
            List<POIXMLDocumentPart> l = new List<POIXMLDocumentPart>();
            foreach(RelationPart rp in relations.Values)
            {
                l.Add(rp.DocumentPart);
            }
            return l.AsReadOnly();
        }

        /// <summary>
        /// Returns the list of child relations for this POIXMLDocumentPart
        /// </summary>
        /// <returns>child relations</returns>

        public IList<RelationPart> RelationParts
        {
            get
            {
                List<RelationPart> l = new List<RelationPart>(relations.Values);
                return l.AsReadOnly();
            }
        }

        /// <summary>
        /// Returns the target <see cref="POIXMLDocumentPart"/>, where a
        /// <see cref="PackageRelationship"/> is set from the <see cref="PackagePart"/> of this
        /// <see cref="POIXMLDocumentPart"/> to the <see cref="PackagePart"/> of the target
        /// <see cref="POIXMLDocumentPart"/> with a <see cref="PackageRelationship.GetId()" />
        /// matching the given parameter value.
        /// </summary>
        /// <param name="id">id
        /// The relation id to look for
        /// </param>
        /// <returns>the target part of the relation, or null, if none exists</returns>

        public POIXMLDocumentPart GetRelationById(String id)
        {
            if(string.IsNullOrEmpty(id) || !relations.ContainsKey(id))
                return null;

            RelationPart rp = relations[id];
            return (rp == null) ? null : rp.DocumentPart;
        }

        /// <summary>
        /// Returns the <see cref="PackageRelationship.GetId()" /> of the
        /// <see cref="PackageRelationship"/>, that sources from the <see cref="PackagePart"/> of
        /// this <see cref="POIXMLDocumentPart"/> to the <see cref="PackagePart"/> of the given
        /// parameter value.
        /// </summary>
        /// <param name="part">part
        /// The <see cref="POIXMLDocumentPart"/> for which the according
        /// relation-id shall be found.
        /// </param>
        /// <returns>The value of the <see cref="PackageRelationship.GetId()" /> or null, if
        /// parts are not related.
        /// </returns>

        public String GetRelationId(POIXMLDocumentPart part)
        {
            foreach(RelationPart rp in relations.Values)
            {
                if(rp.DocumentPart == part)
                {
                    return rp.Relationship.Id;
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
            if(pr == null)
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
            if(relations.ContainsKey(pr.Id))
                relations[pr.Id] = new RelationPart(pr, part);
            else
                relations.Add(pr.Id, new RelationPart(pr, part));
            part.IncrementRelationCounter();
        }

        /// <summary>
        /// Remove the relation to the specified part in this namespace and remove the
        /// part, if it is no longer needed.
        /// </summary>

        protected internal void RemoveRelation(POIXMLDocumentPart part)
        {
            RemoveRelation(part, true);
        }

        /// <summary>
        /// Remove the relation to the specified part in this namespace and remove the
        /// part, if it is no longer needed and flag is set to true.
        /// </summary>
        /// <param name="part">part
        /// The related part, to which the relation shall be Removed.
        /// </param>
        /// <param name="RemoveUnusedParts">RemoveUnusedParts
        /// true, if the part shall be Removed from the namespace if not
        /// needed any longer.
        /// </param>

        protected internal bool RemoveRelation(POIXMLDocumentPart part, bool RemoveUnusedParts)
        {
            String id = GetRelationId(part);
            if(id == null)
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

            if(RemoveUnusedParts)
            {
                /* if last relation to target part was Removed, delete according target part */
                if(part.GetRelationCounter() == 0)
                {
                    try
                    {
                        part.onDocumentRemove();
                    }
                    catch(IOException e)
                    {
                        throw new POIXMLException(e);
                    }
                    GetPackagePart().Package.RemovePart(part.GetPackagePart());
                }
            }
            return true;
        }

        /// <summary>
        /// Returns the parent POIXMLDocumentPart. All parts except root have not-null parent.
        /// </summary>
        /// <returns>the parent POIXMLDocumentPart or <c>null</c> for the root element.</returns>

        public POIXMLDocumentPart GetParent()
        {
            return parent;
        }

        public override String ToString()
        {
            return packagePart == null ? string.Empty : packagePart.ToString();
        }

        /// <summary>
        /// <para>
        /// Save the content in the underlying namespace part.
        /// Default implementation is empty meaning that the namespace part is left unmodified.
        /// </para>
        /// <para>
        /// Sub-classes should override and add logic to marshal the "model" into Ooxml4J.
        /// </para>
        /// <para>
        /// For example, the code saving a generic XML entry may look as follows:
        /// <code><code>
        /// protected void commit()  {
        ///   PackagePart part = GetPackagePart();
        ///   Stream out = part.GetStream();
        ///   XmlObject bean = GetXmlBean(); //the "model" which holds Changes in memory
        ///   bean.save(out, DEFAULT_XML_OPTIONS);
        ///   out.close();
        /// }
        ///  </code></code>
        /// </para>
        /// </summary>

        protected internal virtual void Commit()
        {
        }

        /// <summary>
        /// Save Changes in the underlying OOXML namespace.
        /// Recursively fires <see cref="commit()" /> for each namespace part
        /// </summary>
        /// <param name="alreadySaved">   context set Containing already visited nodes</param>

        protected internal void OnSave(List<PackagePart> alreadySaved)
        {
            // this usually clears out previous content in the part...
            PrepareForCommit();

            Commit();
            alreadySaved.Add(this.GetPackagePart());
            foreach(RelationPart rp in relations.Values)
            {
                POIXMLDocumentPart p = rp.DocumentPart;
                if(!alreadySaved.Contains(p.GetPackagePart()))
                {
                    p.OnSave(alreadySaved);
                }
            }
        }

        /// <summary>
        /// <para>
        /// Ensure that a memory based package part does not have lingering data from previous
        /// commit() calls.
        /// </para>
        /// <para>
        /// Note: This is overwritten for some objects, as *PictureData seem to store the actual content
        /// in the part directly without keeping a copy like all others therefore we need to handle them differently.
        /// </para>
        /// </summary>

        protected internal virtual void PrepareForCommit()
        {
            PackagePart part = this.GetPackagePart();
            if(part != null)
            {
                part.Clear();
            }
        }

        /// <summary>
        /// Create a new child POIXMLDocumentPart
        /// </summary>
        /// <param name="descriptor">the part descriptor</param>
        /// <param name="factory">the factory that will create an instance of the requested relation</param>
        /// <returns>the Created child POIXMLDocumentPart</returns>

        public POIXMLDocumentPart CreateRelationship(POIXMLRelation descriptor, POIXMLFactory factory)
        {
            return CreateRelationship(descriptor, factory, -1, false).DocumentPart;
        }

        public POIXMLDocumentPart CreateRelationship(POIXMLRelation descriptor, POIXMLFactory factory, int idx)
        {
            return CreateRelationship(descriptor, factory, idx, false).DocumentPart;
        }

        /// <summary>
        /// Identifies the next available part number for a part of the given type,
        ///  if possible, otherwise -1 if none are available.
        /// The found (valid) index can then be safely given to
        ///  <see cref="createRelationship(POIXMLRelation, POIXMLFactory, int)" /> or
        ///  <see cref="createRelationship(POIXMLRelation, POIXMLFactory, int, boolean)" />
        ///  without naming clashes.
        /// If parts with other types are already claiming a name for this relationship
        ///  type (eg a <see cref="XSSFRelation.CHART" /> using the drawing part namespace
        ///  normally used by <see cref="XSSFRelation.DRAWINGS" />), those will be considered
        ///  when finding the next spare number.
        /// </summary>
        /// <param name="descriptor">The relationship type to find the part number for</param>
        /// <param name="minIdx">The minimum free index to assign, use -1 for any</param>
        /// <returns>The next free part number, or -1 if none available</returns>

        protected internal int GetNextPartNumber(POIXMLRelation descriptor, int minIdx)
        {
            OPCPackage pkg = packagePart.Package;

            try
            {
                string name = descriptor.DefaultFileName;
                if(name.Equals(descriptor.GetFileName(9999)))
                {
                    // Non-index based, check if default is free
                    PackagePartName ppName = PackagingUriHelper.CreatePartName(name);
                    if(pkg.ContainPart(ppName))
                    {
                        // Default name already taken, not index based, nothing free
                        return -1;
                    }
                    else
                    {
                        // Default name free
                        return 0;
                    }
                }

                // Default to searching from 1, unless they asked for 0+
                int idx = minIdx;
                if(minIdx < 0)
                    idx = 1;
                int maxIdx = minIdx + pkg.GetParts().Count;
                while(idx < maxIdx)
                {
                    name = descriptor.GetFileName(idx);
                    PackagePartName ppName = PackagingUriHelper.CreatePartName(name);
                    if(!pkg.ContainPart(ppName))
                    {
                        return idx;
                    }
                    idx++;
                }
            }
            catch(InvalidFormatException e)
            {
                // Give a general wrapped exception for the problem
                throw new POIXMLException(e);
            }
            return -1;
        }

        /// <summary>
        /// Create a new child POIXMLDocumentPart
        /// </summary>
        /// <param name="descriptor">the part descriptor</param>
        /// <param name="factory">the factory that will create an instance of the requested relation</param>
        /// <param name="idx">part number</param>
        /// <param name="noRelation">if true, then no relationship is Added.</param>
        /// <returns>the Created child POIXMLDocumentPart</returns>

        protected RelationPart CreateRelationship(POIXMLRelation descriptor, POIXMLFactory factory, int idx, bool noRelation)
        {
            try
            {
                PackagePartName ppName = PackagingUriHelper.CreatePartName(descriptor.GetFileName(idx));
                PackageRelationship rel = null;
                PackagePart part = packagePart.Package.CreatePart(ppName, descriptor.ContentType);
                if(!noRelation)
                {
                    /* only add to relations, if according relationship is being Created. */
                    rel = packagePart.AddRelationship(ppName, TargetMode.Internal, descriptor.Relation);
                }
                POIXMLDocumentPart doc = factory.NewDocumentPart(descriptor);
                doc.packagePart = part;
                doc.parent = this;
                if(!noRelation)
                {
                    /* only add to relations, if according relationship is being Created. */
                    AddRelation(rel, doc);
                }
                return new RelationPart(rel, doc);
            }
            catch(PartAlreadyExistsException)
            {
                // Return the specific exception so the user knows
                //  that the name is already taken
                throw;
            }
            catch(Exception e)
            {
                // Give a general wrapped exception for the problem
                throw new POIXMLException(e);
            }
        }

        public TValue PutDictionary<TKey, TValue>(Dictionary<TKey, TValue> dict, TKey key, TValue value)
        {
            TValue oldValue = default(TValue);
            if(dict.ContainsKey(key))
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
            if(dict.ContainsKey(key))
            {
                return dict[key];
            }
            return default(TValue);
        }

        /// <summary>
        /// Iterate through the underlying PackagePart and create child POIXMLFactory instances
        /// using the specified factory
        /// </summary>
        /// <param name="factory">  the factory object that Creates POIXMLFactory instances</param>
        /// <param name="context">  context map Containing already visited noted keyed by tarGetURI</param>

        protected void Read(POIXMLFactory factory, Dictionary<PackagePart, POIXMLDocumentPart> context)
        {
            PackagePart pp = GetPackagePart();
            // add mapping a second time, in case of initial caller hasn't done so
            POIXMLDocumentPart otherChild = PutDictionary(context, pp, this);
            if(otherChild != null && otherChild != this)
            {
                throw new POIXMLException("Unique PackagePart-POIXMLDocumentPart relation broken!");
            }

            if(!pp.HasRelationships)
                return;

            PackageRelationshipCollection rels = packagePart.Relationships;
            List<POIXMLDocumentPart> readLater = new List<POIXMLDocumentPart>();

            // scan breadth-first, so parent-relations are hopefully the shallowest element
            foreach(PackageRelationship rel in rels)
            {
                if(rel.TargetMode == TargetMode.Internal)
                {
                    Uri uri = rel.TargetUri;

                    // check for internal references (e.g. '#Sheet1!A1')
                    PackagePartName relName;
                    //if (uri.getRawFragment() != null)
                    if(uri.OriginalString.IndexOf('#') >= 0)
                    {
                        string path = string.Empty;
                        try
                        {
                            path = uri.AbsolutePath;
                        }
                        catch(InvalidOperationException)
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
                    if(p == null)
                    {
                        //logger.log(POILogger.ERROR, "Skipped invalid entry " + rel.TargetUri);
                        continue;
                    }

                    POIXMLDocumentPart childPart = GetDictionary(context, p);
                    if(childPart == null)
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

            foreach(POIXMLDocumentPart childPart in readLater)
            {
                childPart.Read(factory, context);
            }
        }

        /// <summary>
        /// Get the PackagePart that is the target of a relationship.
        /// </summary>
        /// <param name="rel">The relationship</param>
        /// <returns>The target part</returns>
        /// <exception cref="InvalidFormatException">InvalidFormatException</exception>

        protected PackagePart GetTargetPart(PackageRelationship rel)
        {
            return GetPackagePart().GetRelatedPart(rel);
        }

        /// <summary>
        /// Fired when a new namespace part is Created
        /// </summary>

        internal virtual void OnDocumentCreate()
        {
        }

        /// <summary>
        /// Fired when a namespace part is read
        /// </summary>

        internal virtual void OnDocumentRead()
        {
        }

        /// <summary>
        /// Fired when a namespace part is about to be Removed from the namespace
        /// </summary>

        protected virtual void onDocumentRemove()
        {
        }

        /// <summary>
        /// Retrieves the core document part
        /// </summary>
        /// @since POI 3.14-Beta1

        private static PackagePart GetPartFromOPCPackage(OPCPackage pkg, String coreDocumentRel)
        {
            PackageRelationship coreRel = pkg.GetRelationshipsByType(coreDocumentRel).GetRelationship(0);

            if(coreRel != null)
            {
                PackagePart pp = pkg.GetPart(coreRel);
                if(pp == null)
                {
                    throw new POIXMLException("OOXML file structure broken/invalid - core document '" + coreRel.TargetUri + "' not found.");
                }
                return pp;
            }

            coreRel = pkg.GetRelationshipsByType(PackageRelationshipTypes.STRICT_CORE_DOCUMENT).GetRelationship(0);
            if(coreRel != null)
            {
                throw new POIXMLException("Strict OOXML isn't currently supported, please see bug #57699");
            }

            throw new POIXMLException("OOXML file structure broken/invalid - no core document found!");
        }
    }
}
