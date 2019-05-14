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
        public static CT_SheetPr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SheetPr ctObj = new CT_SheetPr();
            ctObj.syncHorizontal = XmlHelper.ReadBool(node.Attributes[nameof(syncHorizontal)]);
            ctObj.syncVertical = XmlHelper.ReadBool(node.Attributes[nameof(syncVertical)]);
            ctObj.syncRef = XmlHelper.ReadString(node.Attributes[nameof(syncRef)]);
            ctObj.transitionEvaluation = XmlHelper.ReadBool(node.Attributes[nameof(transitionEvaluation)]);
            ctObj.transitionEntry = XmlHelper.ReadBool(node.Attributes[nameof(transitionEntry)]);
            ctObj.published = XmlHelper.ReadBool(node.Attributes[nameof(published)]);
            ctObj.codeName = XmlHelper.ReadString(node.Attributes[nameof(codeName)]);
            ctObj.filterMode = XmlHelper.ReadBool(node.Attributes[nameof(filterMode)]);
            ctObj.enableFormatConditionsCalculation = XmlHelper.ReadBool(node.Attributes[nameof(enableFormatConditionsCalculation)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(tabColor))
                    ctObj.tabColor = CT_Color.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(outlinePr))
                    ctObj.outlinePr = CT_OutlinePr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(pageSetUpPr))
                    ctObj.pageSetUpPr = CT_PageSetUpPr.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(syncHorizontal), this.syncHorizontal, false);
            XmlHelper.WriteAttribute(sw, nameof(syncVertical), this.syncVertical, false);
            XmlHelper.WriteAttribute(sw, nameof(syncRef), this.syncRef);
            XmlHelper.WriteAttribute(sw, nameof(transitionEvaluation), this.transitionEvaluation, false);
            XmlHelper.WriteAttribute(sw, nameof(transitionEntry), this.transitionEntry, false);
            XmlHelper.WriteAttribute(sw, nameof(published), this.published, false);
            XmlHelper.WriteAttribute(sw, nameof(codeName), this.codeName);
            XmlHelper.WriteAttribute(sw, nameof(filterMode), this.filterMode,false);
            XmlHelper.WriteAttribute(sw, nameof(enableFormatConditionsCalculation), this.enableFormatConditionsCalculation, false);
            sw.Write(">");
            this.tabColor?.Write(sw, nameof(tabColor));
            this.outlinePr?.Write(sw, nameof(outlinePr));
            this.pageSetUpPr?.Write(sw, nameof(pageSetUpPr));
            sw.Write($"</{nodeName}>");
        }


        public CT_SheetPr()
        {
            //this.pageSetUpPrField = new CT_PageSetUpPr();
            //this.outlinePrField = new CT_OutlinePr();
            //this.tabColorField = new CT_Color();
            this.syncHorizontal = false;
            this.syncVertical = false;
            this.transitionEvaluation = false;
            this.transitionEntry = false;
            this.published = true;
            this.filterMode = false;
            this.enableFormatConditionsCalculation = true;
        }
        public CT_SheetPr Clone()
        {
            CT_SheetPr newPr = new CT_SheetPr
            {
                codeName = codeName,
                enableFormatConditionsCalculation = enableFormatConditionsCalculation,
                filterMode = filterMode,
                published = published,
                syncHorizontal = syncHorizontal,
                syncRef = syncRef,
                syncVertical = syncVertical,
                transitionEntry = transitionEntry,
                transitionEvaluation = transitionEvaluation
            };

            if (outlinePr != null)
            {
                newPr.outlinePr = outlinePr.Clone();
            }
            if (pageSetUpPr != null)
            {
                newPr.pageSetUpPr = pageSetUpPr.Clone();
            }
            if (tabColor != null)
            {
                newPr.tabColor = tabColor.Copy();
            }
            return newPr;
        }

        public bool IsSetOutlinePr()
        {
            return this.outlinePr != null;
        }
        public bool IsSetPageSetUpPr()
        {
            return this.pageSetUpPr != null;
        }
        public CT_PageSetUpPr AddNewPageSetUpPr()
        {
            this.pageSetUpPr = new CT_PageSetUpPr();
            return this.pageSetUpPr;
        }
        public CT_OutlinePr AddNewOutlinePr()
        {
            this.outlinePr = new CT_OutlinePr();
            return this.outlinePr;
        }


        public CT_Color tabColor { get; set; }

        public CT_OutlinePr outlinePr { get; set; }

        public CT_PageSetUpPr pageSetUpPr { get; set; }

        [DefaultValue(false)]
        public bool syncHorizontal { get; set; }

        [DefaultValue(false)]
        public bool syncVertical { get; set; }

        public string syncRef { get; set; }

        [DefaultValue(false)]
        public bool transitionEvaluation { get; set; }

        [DefaultValue(false)]
        public bool transitionEntry { get; set; }

        [DefaultValue(true)]
        public bool published { get; set; }

        public string codeName { get; set; }

        [DefaultValue(false)]
        public bool filterMode { get; set; }

        [DefaultValue(true)]
        public bool enableFormatConditionsCalculation { get; set; }

        public bool IsSetTabColor()
        {
            return this.tabColor != null;
        }
    }

}
