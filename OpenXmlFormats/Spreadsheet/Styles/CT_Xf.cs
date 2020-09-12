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
    [XmlRoot(ElementName = "xf", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class CT_Xf
    {
        //internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Xf));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });

        private CT_CellAlignment alignmentField = null;

        private CT_CellProtection protectionField = null;

        private CT_ExtensionList extLstField = null;

        private uint numFmtIdField = 0;
        private uint fontIdField = 0;
        private uint fillIdField = 0;
        private uint borderIdField = 0;
        private uint xfIdField = 0;
        private bool quotePrefixField = false;
        private bool pivotButtonField = false;
        private bool applyNumberFormatField = false;
        private bool applyFontField = false;
        private bool applyFillField = false;
        private bool applyBorderField = false;
        private bool applyAlignmentField = false;
        private bool applyProtectionField = false;

        bool numFmtIdSpecifiedField = false;
        bool fontIdSpecifiedField = false;
        bool fillIdSpecifiedField = false;
        bool borderIdSpecifiedField = false;
        bool xfIdSpecifiedField = false;

        public CT_Xf Copy()
        {
            CT_Xf obj = new CT_Xf();
            if (this.alignment!=null)
                obj.alignment = this.alignment.Copy();
            obj.protection = this.protection;
            obj.extLstField = null == extLstField ? null : this.extLstField.Copy();

            obj.applyAlignment = this.applyAlignment;
            obj.applyBorder = this.applyBorder;
            obj.applyFill = this.applyFill;
            obj.applyFont = this.applyFont;
            obj.applyNumberFormat = this.applyNumberFormat;
            obj.applyProtection = this.applyProtection;
            obj.borderId = this.borderId;
            obj.borderIdSpecified = this.borderIdSpecified;
            obj.fillId = this.fillId;
            obj.fillIdSpecifiedField = this.fillIdSpecifiedField;
            obj.fontId = this.fontId;
            obj.fontIdSpecified = this.fontIdSpecified;
            obj.numFmtId = this.numFmtId;
            obj.numFmtIdSpecified = this.numFmtIdSpecified;
            obj.pivotButtonField = this.pivotButtonField;
            obj.quotePrefixField = this.quotePrefixField;
            obj.xfIdField = this.xfIdField;
            return obj;
        }

        public static CT_Xf Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Xf ctObj = new CT_Xf();
            ctObj.numFmtId = XmlHelper.ReadUInt(node.Attributes["numFmtId"]);
            ctObj.fontId = XmlHelper.ReadUInt(node.Attributes["fontId"]);
            ctObj.fillId = XmlHelper.ReadUInt(node.Attributes["fillId"]);
            ctObj.borderId = XmlHelper.ReadUInt(node.Attributes["borderId"]);
            ctObj.xfId = XmlHelper.ReadUInt(node.Attributes["xfId"]);
            ctObj.quotePrefix = XmlHelper.ReadBool(node.Attributes["quotePrefix"]);
            ctObj.pivotButton = XmlHelper.ReadBool(node.Attributes["pivotButton"]);
            ctObj.applyNumberFormat = XmlHelper.ReadBool(node.Attributes["applyNumberFormat"]);
            ctObj.applyFont = XmlHelper.ReadBool(node.Attributes["applyFont"]);
            ctObj.applyFill = XmlHelper.ReadBool(node.Attributes["applyFill"]);
            ctObj.applyBorder = XmlHelper.ReadBool(node.Attributes["applyBorder"]);
            ctObj.applyAlignment = XmlHelper.ReadBool(node.Attributes["applyAlignment"]);
            ctObj.applyProtection = XmlHelper.ReadBool(node.Attributes["applyProtection"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "alignment")
                    ctObj.alignment = CT_CellAlignment.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "protection")
                    ctObj.protection = CT_CellProtection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "numFmtId", this.numFmtId, true);
            XmlHelper.WriteAttribute(sw, "fontId", this.fontId, true);
            XmlHelper.WriteAttribute(sw, "fillId", this.fillId, true);
            XmlHelper.WriteAttribute(sw, "borderId", this.borderId, true);
            XmlHelper.WriteAttribute(sw, "xfId", this.xfId, true);
            XmlHelper.WriteAttribute(sw, "quotePrefix", this.quotePrefix,false);
            XmlHelper.WriteAttribute(sw, "pivotButton", this.pivotButton, false);
            if(this.applyNumberFormat)
                XmlHelper.WriteAttribute(sw, "applyNumberFormat", this.applyNumberFormat);
            XmlHelper.WriteAttribute(sw, "applyFont", this.applyFont, false);
            if (this.applyFill)
                XmlHelper.WriteAttribute(sw, "applyFill", this.applyFill);
            if (this.applyBorder)
                XmlHelper.WriteAttribute(sw, "applyBorder", this.applyBorder, true);
            if (this.applyAlignment)
                XmlHelper.WriteAttribute(sw, "applyAlignment", this.applyAlignment, true);
            if(this.applyProtection)
                XmlHelper.WriteAttribute(sw, "applyProtection", this.applyProtection, true);
            if (this.alignment == null && this.protection == null && this.extLst == null)
            {
                sw.Write("/>");
            }
            else
            {
                sw.Write(">");
                if (this.alignment != null)
                    this.alignment.Write(sw, "alignment");
                if (this.protection != null)
                    this.protection.Write(sw, "protection");
                if (this.extLst != null)
                    this.extLst.Write(sw, "extLst");
                sw.Write(string.Format("</{0}>", nodeName));
            }
        }

        public override string ToString()
        {
            XmlSerializer serializer = new XmlSerializer(typeof(CT_Xf));
            using (StringWriter stream = new StringWriter())
            {
                serializer.Serialize(stream, this);
                return stream.ToString();
            }
        }

        public bool IsSetFontId()
        {
            return this.fontIdSpecifiedField;
        }
        public bool IsSetAlignment()
        {
            return alignmentField != null;
        }
        public void UnsetAlignment()
        {
            this.alignmentField = null;
        }

        public bool IsSetExtLst()
        {
            return this.extLst == null;
        }
        public void UnsetExtLst()
        {
            this.extLst = null;
        }

        public bool IsSetProtection()
        {
            return this.protection != null;
        }
        public void UnsetProtection()
        {
            this.protection = null;
        }
        public bool IsSetLocked()
        {
            // first guess:
            return IsSetProtection() &&  (protectionField.locked == true);
        }
        public CT_CellProtection AddNewProtection()
        {
            this.protectionField = new CT_CellProtection();
            return this.protectionField;
        }

        [XmlElement]
        public CT_CellAlignment alignment
        {
            get { return this.alignmentField; }
            set { this.alignmentField = value; }
        }


        [XmlElement]
        public CT_CellProtection protection
        {
            get { return this.protectionField; }
            set { this.protectionField = value; }
        }

        [XmlElement]
        public CT_ExtensionList extLst
        {
            get { return this.extLstField; }
            set { this.extLstField = value; }
        }
        [XmlAttribute]
        public uint numFmtId
        {
            get { return this.numFmtIdField; }
            set
            {
                this.numFmtIdField = value;
                this.numFmtIdSpecifiedField = true;
            }
        }

        [XmlIgnore]
        public bool numFmtIdSpecified
        {
            get { return numFmtIdSpecifiedField; }
            set { numFmtIdSpecifiedField = value; }
        }

        [XmlAttribute]
        public uint fontId
        {
            get { return this.fontIdField; }
            set
            {
                this.fontIdField = value;
                this.fontIdSpecifiedField = true;
            }
        }

        [XmlIgnore]
        public bool fontIdSpecified
        {
            get { return fontIdSpecifiedField; }
            set { fontIdSpecifiedField = value; }
        }

        [XmlAttribute]
        public uint fillId
        {
            get { return this.fillIdField; }
            set
            {
                this.fillIdField = value;
                this.fillIdSpecifiedField = true;
            }
        }

        [XmlIgnore]
        public bool fillIdSpecified
        {
            get { return fillIdSpecifiedField; }
            set { fillIdSpecifiedField = value; }
        }

        [XmlAttribute]
        public uint borderId
        {
            get { return this.borderIdField; }
            set
            {
                this.borderIdField = value;
                borderIdSpecifiedField = true;
            }
        }
        [XmlIgnore]
        public bool borderIdSpecified
        {
            get { return borderIdSpecifiedField; }
            set { borderIdSpecifiedField = value; }
        }

        [XmlAttribute]
        public uint xfId
        {
            get { return this.xfIdField; }
            set
            {
                this.xfIdField = value;
                this.xfIdSpecifiedField = true;
            }
        }
        [XmlIgnore]
        public bool xfIdSpecified
        {
            get { return xfIdSpecifiedField; }
            set { xfIdSpecifiedField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool quotePrefix
        {
            get { return quotePrefixField; }
            set { this.quotePrefixField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool pivotButton
        {
            get { return pivotButtonField; }
            set { this.pivotButtonField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyNumberFormat
        {
            get { return this.applyNumberFormatField; }
            set
            {
                this.applyNumberFormatField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyFont
        {
            get { return this.applyFontField; }
            set { this.applyFontField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyFill
        {
            get { return this.applyFillField; }
            set { this.applyFillField = value; }
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyBorder
        {
            get { return this.applyBorderField; }
            set { this.applyBorderField = value; }
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyAlignment
        {
            get { return this.applyAlignmentField; }
            set { this.applyAlignmentField = value; }
        }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyProtection
        {
            get { return this.applyProtectionField; }
            set { this.applyProtectionField = value; }
        }

        public bool IsSetApplyFill()
        {
            return this.applyFillField;
        }
    }
}
