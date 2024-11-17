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
    using System;
    using NPOI.POIFS.Common;
    using NPOI.Util;
    using NPOI.OpenXml4Net.Exceptions;
    using System.IO;
    using NPOI.OpenXml4Net.OPC;
    using System.Collections.Generic;
    using NPOI.OpenXml4Net;
    using System.Reflection;
    using NPOI.POIFS.FileSystem;

    public abstract class POIXMLDocument : POIXMLDocumentPart, ICloseable
    {
        public static String DOCUMENT_CREATOR = "NPOI";

        // OLE embeddings relation name
        public static String OLE_OBJECT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";

        // Embedded OPC documents relation name
        public static String PACK_OBJECT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";

        /** The OPC Package */
        private OPCPackage pkg;

        /**
         * The properties of the OPC namespace, opened as needed
         */
        private POIXMLProperties properties;

        protected POIXMLDocument(OPCPackage pkg)
            : base(pkg)
        {
            init(pkg);
        }

        protected POIXMLDocument(OPCPackage pkg, String coreDocumentRel)
            : base(pkg, coreDocumentRel)
        {
            init(pkg);
        }

        private void init(OPCPackage pkg)
        {
            this.pkg = pkg;

        }

        /**
         * Wrapper to open a namespace, returning an IOException
         *  in the event of a problem.
         * Works around shortcomings in java's this() constructor calls
         */
        public static OPCPackage OpenPackage(String path, bool readOnly = false)
        {
            try
            {
                return OPCPackage.Open(path, readOnly ? PackageAccess.READ: PackageAccess.READ_WRITE);
            }
            catch (InvalidFormatException e)
            {
                throw new IOException(e.ToString());
            }
        }

        public OPCPackage Package
        {
            get
            {
                return this.pkg;
            }
        }

        protected PackagePart CorePart
        {
            get
            {
                return GetPackagePart();
            }
        }

        /**
         * Retrieves all the PackageParts which are defined as
         *  relationships of the base document with the
         *  specified content type.
         */
        protected PackagePart[] GetRelatedByType(String contentType)
        {
            PackageRelationshipCollection partsC =
                GetPackagePart().GetRelationshipsByType(contentType);

            PackagePart[] parts = new PackagePart[partsC.Size];
            int count = 0;
            foreach (PackageRelationship rel in partsC)
            {
                parts[count] = GetPackagePart().GetRelatedPart(rel);
                count++;
            }
            return parts;
        }

        /**
         * Checks that the supplied Stream (which MUST
         *  support mark and reSet, or be a PushbackStream)
         *  has a OOXML (zip) header at the start of it.
         * If your Stream does not support mark / reSet,
         *  then wrap it in a PushBackStream, then be
         *  sure to always use that, and not the original!
         * @param inp An Stream which supports either mark/reSet, or is a PushbackStream
         */
        [Obsolete("Use the method from DocumentFactoryHelper")]
        public static bool HasOOXMLHeader(Stream inp)
        {
            return DocumentFactoryHelper.HasOOXMLHeader(inp);
        }

        /**
         * Get the document properties. This gives you access to the
         *  core ooxml properties, and the extended ooxml properties.
         */
        public POIXMLProperties GetProperties()
        {
            if (properties == null)
            {
                try
                {
                    properties = new POIXMLProperties(pkg);
                }
                catch (Exception e)
                {
                    throw new POIXMLException(e);
                }
            }
            return properties;
        }

        /**
         * Get the document's embedded files.
         */
        public abstract List<PackagePart> GetAllEmbedds();

        protected void Load(POIXMLFactory factory)
        {
            Dictionary<PackagePart, POIXMLDocumentPart> context = new Dictionary<PackagePart, POIXMLDocumentPart>();
            try
            {
                Read(factory, context);
            }
            catch (OpenXml4NetException e)
            {
                throw new POIXMLException(e);
            }
            OnDocumentRead();
            context.Clear();
        }
        /**
         * Closes the underlying {@link OPCPackage} from which this
         *  document was read, if there is one
         */
        public void Close()
        {
            if (pkg != null)
            {
                if (pkg.GetPackageAccess() == PackageAccess.READ)
                {
                    pkg.Revert();
                }
                else
                {
                    pkg.Close();
                }
                pkg = null;
            }
        }
        /**
         * Write out this document to an Outputstream.
         *
         * Note - if the Document was opened from a {@link File} rather
         *  than an {@link InputStream}, you <b>must</b> write out to
         *  a different file, overwriting via an OutputStream isn't possible.
         *  
         * If {@code stream} is a {@link java.io.FileOutputStream} on a networked drive
         * or has a high cost/latency associated with each written byte,
         * consider wrapping the OutputStream in a {@link java.io.BufferedOutputStream}
         * to improve write performance.
         * 
         * @param stream - the java Stream you wish to write the file to
         *
         * @exception IOException if anything can't be written.
         */
        public void Write(Stream stream)
        {
            OPCPackage pkg = Package;
            if (pkg == null)
            {
                throw new IOException("Cannot write data, document seems to have been closed already");
            }
            if (!this.GetProperties().CustomProperties.Contains("Generator"))
                this.GetProperties().CustomProperties.AddProperty("Generator", "NPOI");
            if (!this.GetProperties().CustomProperties.Contains("Generator Version"))
                this.GetProperties().CustomProperties.AddProperty("Generator Version", Assembly.GetExecutingAssembly().GetName().Version.ToString(3));
            //force all children to commit their Changes into the underlying OOXML Package
            List<PackagePart> context = new List<PackagePart>();
            OnSave(context);
            context.Clear();

            //save extended and custom properties
            GetProperties().Commit();

            pkg.Save(stream);
        }
    }
}





