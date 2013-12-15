using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;
using System.IO;
using NPOI.OpenXml4Net.Util;
using System.Xml;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_SheetPr
    {

        private CT_Color tabColorField;

        private CT_OutlinePr outlinePrField;

        private CT_PageSetUpPr pageSetUpPrField;

        private bool syncHorizontalField;

        private bool syncVerticalField;

        private string syncRefField;

        private bool transitionEvaluationField;

        private bool transitionEntryField;

        private bool publishedField;

        private string codeNameField;

        private bool filterModeField;

        private bool enableFormatConditionsCalculationField;

        public static CT_SheetPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SheetPr ctObj = new CT_SheetPr();
            ctObj.syncHorizontal = XmlHelper.ReadBool(node.Attributes["syncHorizontal"]);
            ctObj.syncVertical = XmlHelper.ReadBool(node.Attributes["syncVertical"]);
            ctObj.syncRef = XmlHelper.ReadString(node.Attributes["syncRef"]);
            ctObj.transitionEvaluation = XmlHelper.ReadBool(node.Attributes["transitionEvaluation"]);
            ctObj.transitionEntry = XmlHelper.ReadBool(node.Attributes["transitionEntry"]);
            ctObj.published = XmlHelper.ReadBool(node.Attributes["published"]);
            ctObj.codeName = XmlHelper.ReadString(node.Attributes["codeName"]);
            ctObj.filterMode = XmlHelper.ReadBool(node.Attributes["filterMode"]);
            ctObj.enableFormatConditionsCalculation = XmlHelper.ReadBool(node.Attributes["enableFormatConditionsCalculation"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tabColor")
                    ctObj.tabColor = CT_Color.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "outlinePr")
                    ctObj.outlinePr = CT_OutlinePr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pageSetUpPr")
                    ctObj.pageSetUpPr = CT_PageSetUpPr.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "syncHorizontal", this.syncHorizontal, false);
            XmlHelper.WriteAttribute(sw, "syncVertical", this.syncVertical, false);
            XmlHelper.WriteAttribute(sw, "syncRef", this.syncRef);
            XmlHelper.WriteAttribute(sw, "transitionEvaluation", this.transitionEvaluation, false);
            XmlHelper.WriteAttribute(sw, "transitionEntry", this.transitionEntry, false);
            XmlHelper.WriteAttribute(sw, "published", this.published, false);
            XmlHelper.WriteAttribute(sw, "codeName", this.codeName);
            XmlHelper.WriteAttribute(sw, "filterMode", this.filterMode,false);
            XmlHelper.WriteAttribute(sw, "enableFormatConditionsCalculation", this.enableFormatConditionsCalculation, false);
            sw.Write(">");
            if (this.tabColor != null)
                this.tabColor.Write(sw, "tabColor");
            if (this.outlinePr != null)
                this.outlinePr.Write(sw, "outlinePr");
            if (this.pageSetUpPr != null)
                this.pageSetUpPr.Write(sw, "pageSetUpPr");
            sw.Write(string.Format("</{0}>", nodeName));
        }


        public CT_SheetPr()
        {
            //this.pageSetUpPrField = new CT_PageSetUpPr();
            //this.outlinePrField = new CT_OutlinePr();
            //this.tabColorField = new CT_Color();
            this.syncHorizontalField = false;
            this.syncVerticalField = false;
            this.transitionEvaluationField = false;
            this.transitionEntryField = false;
            this.publishedField = true;
            this.filterModeField = false;
            this.enableFormatConditionsCalculationField = true;
        }
        public bool IsSetOutlinePr()
        {
            return this.outlinePrField != null;
        }
        public bool IsSetPageSetUpPr()
        {
            return this.pageSetUpPrField != null;
        }
        public CT_PageSetUpPr AddNewPageSetUpPr()
        {
            this.pageSetUpPrField = new CT_PageSetUpPr();
            return this.pageSetUpPrField;
        }
        public CT_OutlinePr AddNewOutlinePr()
        {
            this.outlinePrField = new CT_OutlinePr();
            return this.outlinePrField;
        }


        public CT_Color tabColor
        {
            get
            {
                return this.tabColorField;
            }
            set
            {
                this.tabColorField = value;
            }
        }

        public CT_OutlinePr outlinePr
        {
            get
            {
                return this.outlinePrField;
            }
            set
            {
                this.outlinePrField = value;
            }
        }

        public CT_PageSetUpPr pageSetUpPr
        {
            get
            {
                return this.pageSetUpPrField;
            }
            set
            {
                this.pageSetUpPrField = value;
            }
        }

        [DefaultValue(false)]
        public bool syncHorizontal
        {
            get
            {
                return this.syncHorizontalField;
            }
            set
            {
                this.syncHorizontalField = value;
            }
        }

        [DefaultValue(false)]
        public bool syncVertical
        {
            get
            {
                return this.syncVerticalField;
            }
            set
            {
                this.syncVerticalField = value;
            }
        }

        public string syncRef
        {
            get
            {
                return this.syncRefField;
            }
            set
            {
                this.syncRefField = value;
            }
        }

        [DefaultValue(false)]
        public bool transitionEvaluation
        {
            get
            {
                return this.transitionEvaluationField;
            }
            set
            {
                this.transitionEvaluationField = value;
            }
        }

        [DefaultValue(false)]
        public bool transitionEntry
        {
            get
            {
                return this.transitionEntryField;
            }
            set
            {
                this.transitionEntryField = value;
            }
        }

        [DefaultValue(true)]
        public bool published
        {
            get
            {
                return this.publishedField;
            }
            set
            {
                this.publishedField = value;
            }
        }

        public string codeName
        {
            get
            {
                return this.codeNameField;
            }
            set
            {
                this.codeNameField = value;
            }
        }

        [DefaultValue(false)]
        public bool filterMode
        {
            get
            {
                return this.filterModeField;
            }
            set
            {
                this.filterModeField = value;
            }
        }

        [DefaultValue(true)]
        public bool enableFormatConditionsCalculation
        {
            get
            {
                return this.enableFormatConditionsCalculationField;
            }
            set
            {
                this.enableFormatConditionsCalculationField = value;
            }
        }

        public bool IsSetTabColor()
        {
            return this.tabColor != null;
        }
    }

}
