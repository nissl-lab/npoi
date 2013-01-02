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

using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.Exceptions;
using System;
using System.IO;
using NPOI.Util;
namespace NPOI.Util
{

    /**
     * Provides handy methods to work with OOXML namespaces
     *
     * @author Yegor Kozlov
     */
    public class PackageHelper
    {

        public static OPCPackage Open(Stream is1)
        {
            try
            {
                return OPCPackage.Open(is1);
            }
            catch (InvalidFormatException e)
            {
                throw new POIXMLException(e);
            }
        }

        /**
         * Clone the specified namespace.
         *
         * @param   pkg   the namespace to clone
         * @param   file  the destination file
         * @return  the Cloned namespace
         */
        public static OPCPackage Clone(OPCPackage pkg, string path)
        {

            OPCPackage dest = OPCPackage.Create(path);
            PackageRelationshipCollection rels = pkg.Relationships;
            foreach (PackageRelationship rel in rels)
            {
                PackagePart part = pkg.GetPart(rel);
                PackagePart part_tgt;
                if (rel.RelationshipType.Equals(PackageRelationshipTypes.CORE_PROPERTIES))
                {
                    CopyProperties(pkg.GetPackageProperties(), dest.GetPackageProperties());
                    continue;
                }
                dest.AddRelationship(part.PartName, (TargetMode)rel.TargetMode, rel.RelationshipType);
                part_tgt = dest.CreatePart(part.PartName, part.ContentType);

                Stream out1 = part_tgt.GetOutputStream();
                IOUtils.Copy(part.GetInputStream(), out1);
                out1.Close();

                if (part.HasRelationships)
                {
                    Copy(pkg, part, dest, part_tgt);
                }
            }
            dest.Close();

            //the temp file will be deleted when JVM terminates
            //new File(path).deleteOnExit();
            return OPCPackage.Open(path);
        }

        /**
         * Creates an empty file in the default temporary-file directory,
         */
        public string CreateTempFile()
        {
            string file = TempFile.GetTempFilePath("poi-ooxml-", ".tmp");
            return file;
        }

        /**
         * Recursively copy namespace parts to the destination namespace
         */
        private static void Copy(OPCPackage pkg, PackagePart part, OPCPackage tgt, PackagePart part_tgt) {
        PackageRelationshipCollection rels = part.Relationships;
        if(rels != null) 
            foreach (PackageRelationship rel in rels) {
            PackagePart p;
            if(rel.TargetMode == TargetMode.External){
                part_tgt.AddExternalRelationship(rel.TargetUri.ToString(), rel.RelationshipType, rel.Id);
                //external relations don't have associated namespace parts
                continue;
            }
            Uri uri = rel.TargetUri;

            if(uri.Fragment != null) {
                part_tgt.AddRelationship(uri, (TargetMode)rel.TargetMode, rel.RelationshipType, rel.Id);
                continue;
            }
            PackagePartName relName = PackagingUriHelper.CreatePartName(rel.TargetUri);
            p = pkg.GetPart(relName);
            part_tgt.AddRelationship(p.PartName, (TargetMode)rel.TargetMode, rel.RelationshipType, rel.Id);




            PackagePart dest;
            if(!tgt.ContainPart(p.PartName)){
                dest = tgt.CreatePart(p.PartName, p.ContentType);
                Stream out1 = dest.GetOutputStream();
                IOUtils.Copy(p.GetInputStream(), out1);
                out1.Close();
                Copy(pkg, p, tgt, dest);
            }
        }
    }

        /**
         * Copy core namespace properties
         *
         * @param src source properties
         * @param tgt target properties
         */
        private static void CopyProperties(PackageProperties src, PackageProperties tgt)
        {
            tgt.SetCategoryProperty(src.GetCategoryProperty());
            tgt.SetContentStatusProperty(src.GetContentStatusProperty());
            tgt.SetContentTypeProperty(src.GetContentTypeProperty());
            tgt.SetCreatorProperty(src.GetCreatorProperty());
            tgt.SetDescriptionProperty(src.GetDescriptionProperty());
            tgt.SetIdentifierProperty(src.GetIdentifierProperty());
            tgt.SetKeywordsProperty(src.GetKeywordsProperty());
            tgt.SetLanguageProperty(src.GetLanguageProperty());
            tgt.SetRevisionProperty(src.GetRevisionProperty());
            tgt.SetSubjectProperty(src.GetSubjectProperty());
            tgt.SetTitleProperty(src.GetTitleProperty());
            tgt.SetVersionProperty(src.GetVersionProperty());
        }
    }


}