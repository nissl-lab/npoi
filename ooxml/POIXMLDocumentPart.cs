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

 
        private PackagePart packagePart;
        private PackageRelationship packageRel;
        private POIXMLDocumentPart parent;
        private Dictionary<String, POIXMLDocumentPart> relations = new Dictionary<String, POIXMLDocumentPart>();

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
        {
            PackageRelationship coreRel = pkg.GetRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT).GetRelationship(0);
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
         * Creates an POIXMLDocumentPart representing the given namespace part and relationship.
         * Called by {@link #read(POIXMLFactory, java.util.Map)} when Reading in an exisiting file.
         *
         * @param part - The namespace part that holds xml data represenring this sheet.
         * @param rel - the relationship of the given namespace part
         * @see #read(POIXMLFactory, java.util.Map) 
         */
        public POIXMLDocumentPart(PackagePart part, PackageRelationship rel)
        {
            this.packagePart = part;
            this.packageRel = rel;
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
        public POIXMLDocumentPart(POIXMLDocumentPart parent, PackagePart part, PackageRelationship rel)
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
                packagePart.GetRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT);
            if (cores.Size != 1)
            {
                throw new InvalidOperationException(
                    "Tried to rebase using " + PackageRelationshipTypes.CORE_DOCUMENT +
                    " but found " + cores.Size + " parts of the right type"
                );
            }
            packageRel = cores.GetRelationship(0);
            packagePart = packagePart.GetRelatedPart(packageRel);
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
        public PackageRelationship GetPackageRelationship()
        {
            return packageRel;
        }

        /**
         * Returns the list of child relations for this POIXMLDocumentPart
         *
         * @return child relations
         */
        public List<POIXMLDocumentPart> GetRelations()
        {
            return new List<POIXMLDocumentPart>(relations.Values);
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
            if (string.IsNullOrEmpty(id))
                return null;
            return relations[id];
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
            foreach (KeyValuePair<String, POIXMLDocumentPart> entry in relations)
            {
                if (entry.Value == part)
                {
                    return entry.Key;
                }
            }
            return null;
        }

        /**
         * Add a new child POIXMLDocumentPart
         *
         * @param part the child to add
         */
        public void AddRelation(String id, POIXMLDocumentPart part)
        {
            relations[id] = part;
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
            foreach (POIXMLDocumentPart p in relations.Values)
            {
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
            return CreateRelationship(descriptor, factory, -1, false);
        }

        public POIXMLDocumentPart CreateRelationship(POIXMLRelation descriptor, POIXMLFactory factory, int idx)
        {
            return CreateRelationship(descriptor, factory, idx, false);
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
        protected POIXMLDocumentPart CreateRelationship(POIXMLRelation descriptor, POIXMLFactory factory, int idx, bool noRelation)
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
                POIXMLDocumentPart doc = factory.CreateDocumentPart(descriptor);
                doc.packageRel = rel;
                doc.packagePart = part;
                doc.parent = this;
                if (!noRelation)
                {
                    /* only add to relations, if according relationship is being Created. */
                    AddRelation(rel.Id, doc);
                }
                return doc;
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

        /**
         * Iterate through the underlying PackagePart and create child POIXMLFactory instances
         * using the specified factory
         *
         * @param factory   the factory object that Creates POIXMLFactory instances
         * @param context   context map Containing already visited noted keyed by tarGetURI
         */
        protected void Read(POIXMLFactory factory, Dictionary<PackagePart, POIXMLDocumentPart> context)
        {
            try
            {
                PackageRelationshipCollection rels = packagePart.Relationships;
                foreach (PackageRelationship rel in rels)
                {
                    if (rel.TargetMode == TargetMode.Internal)
                    {
                        Uri uri = rel.TargetUri;

                        PackagePart p;
                        
                        if (uri.OriginalString.IndexOf('#')>=0)
                        {
                            /*
                             * For internal references (e.g. '#Sheet1!A1') the namespace part is null
                             */
                            p = null;
                        }
                        else
                        {
                            PackagePartName relName = PackagingUriHelper.CreatePartName(uri);
                            p = packagePart.Package.GetPart(relName);
                            if (p == null)
                            {
                                logger.Log(POILogger.ERROR, "Skipped invalid entry " + rel.TargetUri);
                                continue;
                            }
                        }

                        if (p == null || !context.ContainsKey(p))
                        {
                            POIXMLDocumentPart childPart = factory.CreateDocumentPart(this, rel, p);
                            childPart.parent = this;
                            AddRelation(rel.Id, childPart);
                            if (p != null)
                            {
                                context[p] = childPart;
                                if (p.HasRelationships) childPart.Read(factory, context);
                            }
                        }
                        else
                        {
                            AddRelation(rel.Id, context[p]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if ((null != ex.InnerException) && (null != ex.InnerException.InnerException))
                {
                    // this type of exception is thrown when the XML Serialization does not match the input.
                    logger.Log(1, ex.InnerException.InnerException);
                }
                throw;
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
    }
}






