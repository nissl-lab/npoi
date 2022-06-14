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
using NPOI.OpenXml4Net.Exceptions;
using NPOI.OpenXml4Net.OPC;
using NPOI.Util;
using System;

namespace NPOI
{


    /**
     * Defines a factory API that enables sub-classes to create instances of <code>POIXMLDocumentPart</code>
     *
     * @author Yegor Kozlov
     */
    public abstract class POIXMLFactory
    {
        private static POILogger LOGGER = POILogFactory.GetLogger(typeof(POIXMLFactory));

        private static Type[] PARENT_PART = { typeof(POIXMLDocumentPart), typeof(PackagePart) };
        private static Type[] ORPHAN_PART = { typeof(PackagePart) };

        /**
         * Create a POIXMLDocumentPart from existing package part and relation. This method is called
         * from {@link POIXMLDocument#load(POIXMLFactory)} when parsing a document
         *
         * @param parent parent part
         * @param rel   the package part relationship
         * @param part  the PackagePart representing the created instance
         * @return A new instance of a POIXMLDocumentPart.
         * 
         * @since by POI 3.14-Beta1
         */
        public virtual POIXMLDocumentPart CreateDocumentPart(POIXMLDocumentPart parent, PackagePart part)
        {
            PackageRelationship rel = GetPackageRelationship(parent, part);
            POIXMLRelation descriptor = GetDescriptor(rel.RelationshipType);

            if (descriptor == null || descriptor.RelationClass == null)
            {
                LOGGER.Log(POILogger.DEBUG, "using default POIXMLDocumentPart for " + rel.RelationshipType);
                return new POIXMLDocumentPart(parent, part);
            }
            Type cls = descriptor.RelationClass;
            try
            {
                try
                {
                    return CreateDocumentPart(cls, PARENT_PART, new Object[] { parent, part });
                }
                catch (MissingMethodException)
                {
                    return CreateDocumentPart(cls, ORPHAN_PART, new Object[] { part });
                }
            }
            catch (Exception e)
            {
                throw new POIXMLException(e);
            }
        }

        /**
         * Need to delegate instantiation to sub class because of constructor visibility
         * @since POI 3.14-Beta1
         */
        protected abstract POIXMLDocumentPart CreateDocumentPart(Type cls, Type[] classes, Object[] values);

        /**
         * returns the descriptor for the given relationship type 
         *
         * @return the descriptor or null if type is unknown
         * 
         * @since POI 3.14-Beta1
         */
        protected abstract POIXMLRelation GetDescriptor(String relationshipType);

        /**
         * Create a POIXMLDocumentPart from existing package part and relation. This method is called
         * from {@link POIXMLDocument#load(POIXMLFactory)} when parsing a document
         *
         * @param parent parent part
         * @param rel   the package part relationship
         * @param part  the PackagePart representing the created instance
         * @return A new instance of a POIXMLDocumentPart.
         * 
         * @deprecated in POI 3.14, scheduled for removal in POI 3.16
         */
        [Obsolete("deprecated in POI 3.14, scheduled for removal in POI 3.16")]
        public virtual POIXMLDocumentPart CreateDocumentPart(POIXMLDocumentPart parent, PackageRelationship rel, PackagePart part)
        {
            return CreateDocumentPart(parent, part);
        }

        /**
         * Create a new POIXMLDocumentPart using the supplied descriptor. This method is used when adding new parts
         * to a document, for example, when adding a sheet to a workbook, slide to a presentation, etc.
         *
         * @param descriptor  describes the object to create
         * @return A new instance of a POIXMLDocumentPart.
         */
        public POIXMLDocumentPart NewDocumentPart(POIXMLRelation descriptor)
        {
            Type cls = descriptor.RelationClass;
            try
            {
                return CreateDocumentPart(cls, null, null);
            }
            catch (Exception e)
            {
                throw new POIXMLException(e);
            }
        }

        /**
         * Retrieves the package relationship of the child part within the parent
         * 
         * @since POI 3.14-Beta1
         */
        protected PackageRelationship GetPackageRelationship(POIXMLDocumentPart parent, PackagePart part)
        {
            try
            {
                String partName = part.PartName.Name;
                foreach (PackageRelationship pr in parent.GetPackagePart().Relationships)
                {
                    String packName = pr.TargetUri.OriginalString;// toASCIIString();
                    if (packName.Equals(partName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        return pr;
                    }
                }
            }
            catch (InvalidFormatException e)
            {
                throw new POIXMLException("error while determining package relations", e);
            }

            throw new POIXMLException("package part isn't a child of the parent document.");
        }
    }


}




