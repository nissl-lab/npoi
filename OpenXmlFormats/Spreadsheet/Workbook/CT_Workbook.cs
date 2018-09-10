using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "workbook",
        IsNullable = false)]
    public class CT_Workbook
    {
        // all elements optional except sheets, only fileRecoveryPr may be listed more than once
        private CT_FileVersion fileVersionField;

        private CT_FileSharing fileSharingField;

        private CT_WorkbookPr workbookPrField = null;

        private CT_WorkbookProtection workbookProtectionField;

        private CT_BookViews bookViewsField = null;

        private CT_Sheets sheetsField = new CT_Sheets();

        private CT_FunctionGroups functionGroupsField;

        private CT_ExternalReferences externalReferencesField;

        private CT_DefinedNames definedNamesField = null;

        private CT_CalcPr calcPrField = null;

        private CT_OleSize oleSizeField;

        private CT_CustomWorkbookViews customWorkbookViewsField;

        private CT_PivotCaches pivotCachesField;

        private CT_SmartTagPr smartTagPrField;

        private CT_SmartTagTypes smartTagTypesField;

        private CT_WebPublishing webPublishingField;

        private List<CT_FileRecoveryPr> fileRecoveryPrField = null;

        private CT_WebPublishObjects webPublishObjectsField;

        private CT_ExtensionList extLstField;

        public CT_Workbook()
        {
        }
        public static CT_Workbook Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Workbook ctObj = new CT_Workbook();
            ctObj.fileRecoveryPr = new List<CT_FileRecoveryPr>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "fileVersion")
                    ctObj.fileVersion = CT_FileVersion.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fileSharing")
                    ctObj.fileSharing = CT_FileSharing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "workbookPr")
                    ctObj.workbookPr = CT_WorkbookPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "workbookProtection")
                    ctObj.workbookProtection = CT_WorkbookProtection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bookViews")
                    ctObj.bookViews = CT_BookViews.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sheets")
                    ctObj.sheets = CT_Sheets.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "functionGroups")
                    ctObj.functionGroups = CT_FunctionGroups.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "externalReferences")
                    ctObj.externalReferences = CT_ExternalReferences.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "definedNames")
                    ctObj.definedNames = CT_DefinedNames.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "calcPr")
                    ctObj.calcPr = CT_CalcPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "oleSize")
                    ctObj.oleSize = CT_OleSize.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "customWorkbookViews")
                    ctObj.customWorkbookViews = CT_CustomWorkbookViews.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pivotCaches")
                    ctObj.pivotCaches = CT_PivotCaches.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "smartTagPr")
                    ctObj.smartTagPr = CT_SmartTagPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "smartTagTypes")
                    ctObj.smartTagTypes = CT_SmartTagTypes.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "webPublishing")
                    ctObj.webPublishing = CT_WebPublishing.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "webPublishObjects")
                    ctObj.webPublishObjects = CT_WebPublishObjects.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fileRecoveryPr")
                    ctObj.fileRecoveryPr.Add(CT_FileRecoveryPr.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            if (this.fileVersion != null)
                this.fileVersion.Write(sw, "fileVersion");
            if (this.fileSharing != null)
                this.fileSharing.Write(sw, "fileSharing");
            if (this.workbookPr != null)
                this.workbookPr.Write(sw, "workbookPr");
            if (this.workbookProtection != null)
                this.workbookProtection.Write(sw, "workbookProtection");
            if (this.bookViews != null)
                this.bookViews.Write(sw, "bookViews");
            if (this.sheets != null)
                this.sheets.Write(sw, "sheets");
            if (this.functionGroups != null)
                this.functionGroups.Write(sw, "functionGroups");
            if (this.externalReferences != null)
                this.externalReferences.Write(sw, "externalReferences");
            if (this.definedNames != null)
                this.definedNames.Write(sw, "definedNames");
            if (this.calcPr != null)
                this.calcPr.Write(sw, "calcPr");
            if (this.oleSize != null)
                this.oleSize.Write(sw, "oleSize");
            if (this.customWorkbookViews != null)
                this.customWorkbookViews.Write(sw, "customWorkbookViews");
            if (this.pivotCaches != null)
                this.pivotCaches.Write(sw, "pivotCaches");
            if (this.smartTagPr != null)
                this.smartTagPr.Write(sw, "smartTagPr");
            if (this.smartTagTypes != null)
                this.smartTagTypes.Write(sw, "smartTagTypes");
            if (this.webPublishing != null)
                this.webPublishing.Write(sw, "webPublishing");
            if (this.webPublishObjects != null)
                this.webPublishObjects.Write(sw, "webPublishObjects");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            if (this.fileRecoveryPr != null)
            {
                foreach (CT_FileRecoveryPr x in this.fileRecoveryPr)
                {
                    x.Write(sw, "fileRecoveryPr");
                }
            }
            sw.Write("</workbook>");
        }


        public CT_WorkbookPr AddNewWorkbookPr()
        {
            this.workbookPrField = new CT_WorkbookPr();
            return this.workbookPrField;
        }

        public CT_CalcPr AddNewCalcPr()
        {
            this.calcPrField = new CT_CalcPr();
            return this.calcPrField;
        }
        public CT_Sheets AddNewSheets()
        {
            this.sheetsField = new CT_Sheets();
            return this.sheetsField;
        }
        public CT_BookViews AddNewBookViews()
        {
            this.bookViewsField = new CT_BookViews();
            return this.bookViewsField;
        }
        public bool IsSetWorkbookPr()
        {
            return this.workbookPrField != null;
        }
        public bool IsSetCalcPr()
        {
            return this.calcPrField != null;
        }
        public bool IsSetSheets()
        {
            return this.sheetsField != null;
        }
        public bool IsSetBookViews()
        {
            return this.bookViewsField != null;
        }

        public bool IsSetExternalReferences()
        {
            return this.externalReferencesField != null;
        }

        public bool IsSetDefinedNames()
        {
            return this.definedNamesField != null;
        }
        public CT_DefinedNames AddNewDefinedNames()
        {
            this.definedNamesField = new CT_DefinedNames();
            return this.definedNamesField;
        }
        //AddNewWorkbookView()
        public void SetDefinedNames(CT_DefinedNames definedNames)
        {
            this.definedNamesField = definedNames;
        }
        public void unsetDefinedNames()
        {
            this.definedNamesField = null;
        }
        [XmlElement]
        public CT_FileVersion fileVersion
        {
            get
            {
                return this.fileVersionField;
            }
            set
            {
                this.fileVersionField = value;
            }
        }
        [XmlElement]
        public CT_FileSharing fileSharing
        {
            get
            {
                return this.fileSharingField;
            }
            set
            {
                this.fileSharingField = value;
            }
        }

        [XmlElement]
        public CT_WorkbookPr workbookPr
        {
            get
            {
                return this.workbookPrField;
            }
            set
            {
                this.workbookPrField = value;
            }
        }

        [XmlElement]
        public CT_WorkbookProtection workbookProtection
        {
            get
            {
                return this.workbookProtectionField;
            }
            set
            {
                this.workbookProtectionField = value;
            }
        }

        [XmlElement("bookViews", IsNullable = false)]
        public CT_BookViews bookViews
        {
            get
            {
                return this.bookViewsField;
            }
            set
            {
                this.bookViewsField = value;
            }
        }


        [XmlElement("sheets", IsNullable = false)]
        public CT_Sheets sheets
        {
            get
            {
                return this.sheetsField;
            }
            set
            {
                this.sheetsField = value;
            }
        }

        [XmlElement]
        public CT_FunctionGroups functionGroups
        {
            get
            {
                return this.functionGroupsField;
            }
            set
            {
                this.functionGroupsField = value;
            }
        }

        //[XmlArray(Order = 7)]
        //[XmlArrayItem("externalReference", IsNullable = false)]
        [XmlElement]
        public CT_ExternalReferences externalReferences
        {
            get
            {
                return this.externalReferencesField;
            }
            set
            {
                this.externalReferencesField = value;
            }
        }

        //[XmlArray(Order = 8)]
        //[XmlArrayItem("definedName", IsNullable = false)]
        [XmlElement]
        public CT_DefinedNames definedNames
        {
            get
            {
                return this.definedNamesField;
            }
            set
            {
                this.definedNamesField = value;
            }
        }

        [XmlElement]
        public CT_CalcPr calcPr
        {
            get
            {
                return this.calcPrField;
            }
            set
            {
                this.calcPrField = value;
            }
        }

        [XmlElement]
        public CT_OleSize oleSize
        {
            get
            {
                return this.oleSizeField;
            }
            set
            {
                this.oleSizeField = value;
            }
        }

        //[XmlArray(Order = 11)]
        //[XmlArrayItem("customWorkbookView", IsNullable = false)]
        [XmlElement]
        public CT_CustomWorkbookViews customWorkbookViews
        {
            get
            {
                return this.customWorkbookViewsField;
            }
            set
            {
                this.customWorkbookViewsField = value;
            }
        }

        //[XmlArray(Order = 12)]
        //[XmlArrayItem("pivotCache", IsNullable = false)]
        [XmlElement]
        public CT_PivotCaches pivotCaches
        {
            get
            {
                return this.pivotCachesField;
            }
            set
            {
                this.pivotCachesField = value;
            }
        }

        [XmlElement]
        public CT_SmartTagPr smartTagPr
        {
            get
            {
                return this.smartTagPrField;
            }
            set
            {
                this.smartTagPrField = value;
            }
        }

        //[XmlArray(Order = 14)]
        //[XmlArrayItem("smartTagType", IsNullable = false)]
        [XmlElement]
        public CT_SmartTagTypes smartTagTypes
        {
            get
            {
                return this.smartTagTypesField;
            }
            set
            {
                this.smartTagTypesField = value;
            }
        }

        [XmlElement]
        public CT_WebPublishing webPublishing
        {
            get
            {
                return this.webPublishingField;
            }
            set
            {
                this.webPublishingField = value;
            }
        }

        [XmlElement]
        public List<CT_FileRecoveryPr> fileRecoveryPr
        {
            get
            {
                return this.fileRecoveryPrField;
            }
            set
            {
                this.fileRecoveryPrField = value;
            }
        }

        [XmlElement]
        public CT_WebPublishObjects webPublishObjects
        {
            get
            {
                return this.webPublishObjectsField;
            }
            set
            {
                this.webPublishObjectsField = value;
            }
        }

        [XmlElement]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        public bool IsSetPivotCaches()
        {
            return this.pivotCachesField != null;
        }

        public CT_PivotCaches AddNewPivotCaches()
        {
            this.pivotCachesField = new CT_PivotCaches();
            return this.pivotCachesField;
        }
    }
}
