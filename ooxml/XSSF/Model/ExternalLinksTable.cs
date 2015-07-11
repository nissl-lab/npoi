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
namespace NPOI.XSSF.Model
{
    using System;
    using NPOI.SS.UserModel;
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.OpenXml4Net.OPC;
    using System.IO;
    using System.Xml;
    using System.Collections.Generic;
    using NPOI.OpenXmlFormats.Spreadsheet.Document;


    /**
     * Holds details of links to parts of other workbooks (eg named ranges),
     *  along with the most recently seen values for what they point to.
     */
    public class ExternalLinksTable : POIXMLDocumentPart
    {
        private CT_ExternalLink link;

        public ExternalLinksTable()
            : base()
        {
            link = new CT_ExternalLink();
            link.AddNewExternalBook();
        }

        internal ExternalLinksTable(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
            ;
            ReadFrom(part.GetInputStream());
        }

        public void ReadFrom(Stream is1)
        {
            try
            {
                XmlDocument xmldoc = ConvertStreamToXml(is1);
                ExternalLinkDocument doc = ExternalLinkDocument.Parse(xmldoc, NamespaceManager);
                link = doc.ExternalLink;
            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }
        public void WriteTo(Stream out1)
        {
            ExternalLinkDocument doc = new ExternalLinkDocument();
            doc.ExternalLink = (/*setter*/link);
            doc.Save(out1);
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            WriteTo(out1);
            out1.Close();
        }

        /**
         * Returns the underlying xmlbeans object for the external
         *  link table
         */
        public CT_ExternalLink CTExternalLink
        {
            get
            {
                return link;
            }
        }

        /**
         * get or set the last recorded name of the file that this
         *  is linked to
         */
        public virtual String LinkedFileName
        {
            get
            {
                String rId = link.externalBook.id;
                PackageRelationship rel = GetPackagePart().GetRelationship(rId);
                if (rel != null && rel.TargetMode == TargetMode.External)
                {
                    return rel.TargetUri.ToString();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                String rId = link.externalBook.id;

                if (string.IsNullOrEmpty(rId))
                {
                    // We're a new External Link Table, so nothing to remove
                }
                else
                {
                    // Relationships can't be changed, so remove the old one
                    GetPackagePart().RemoveRelationship(rId);
                }

                // Have a new one added
                PackageRelationship newRel = GetPackagePart().AddExternalRelationship(
                                        value, PackageRelationshipTypes.EXTERNAL_LINK_PATH);
                link.externalBook.id = (newRel.Id);
            }
        }


        public List<String> SheetNames
        {
            get
            {
                CT_ExternalSheetName[] sheetNames =
                        link.externalBook.sheetNames.sheetName;
                List<String> names = new List<String>(sheetNames.Length);
                foreach (CT_ExternalSheetName name in sheetNames)
                {
                    names.Add(name.val);
                }
                return names;
            }
        }

        public List<IName> DefinedNames
        {
            get
            {
                CT_ExternalDefinedName[] extNames =
                        link.externalBook.definedNames.definedName;
                List<IName> names = new List<IName>(extNames.Length);
                foreach (CT_ExternalDefinedName extName in extNames)
                {
                    names.Add(new ExternalName(extName, this));
                }
                return names;
            }
        }


        // TODO Last seen data


        protected internal class ExternalName : IName
        {
            private ExternalLinksTable externalLinkTable;
            private CT_ExternalDefinedName name;
            protected internal ExternalName(CT_ExternalDefinedName name, ExternalLinksTable externalLinkTable)
            {
                this.name = name;
                this.externalLinkTable = externalLinkTable;
            }

            public String NameName
            {
                get
                {
                    return name.name;
                }
                set
                {
                    this.name.name = (/*setter*/value);
                }
            }

            public String SheetName
            {
                get
                {
                    int sheetId = SheetIndex;
                    if (sheetId >= 0)
                    {
                        return externalLinkTable.SheetNames[(sheetId)];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            public int SheetIndex
            {
                get
                {
                    if (name.IsSetSheetId())
                    {
                        return (int)name.sheetId;
                    }
                    return -1;
                }
                set
                {
                    name.sheetId = (uint)value;
                }
            }
            public String RefersToFormula
            {
                get
                {
                    // Return, without the leading =
                    return name.refersTo.Substring(1);
                }
                set
                {
                    // Save with leading =
                    name.refersTo = (/*setter*/'=' + value);
                }
            }
            public bool IsFunctionName
            {
                get
                {
                    return false;
                }
            }
            public bool IsDeleted
            {
                get
                {
                    return false;
                }
            }

            public String Comment
            {
                get
                {
                    return null;
                }
                set
                {
                    throw new InvalidOperationException("Not Supported");
                }
            }
            public void SetFunction(bool value)
            {
                throw new InvalidOperationException("Not Supported");
            }
        }
    }
}