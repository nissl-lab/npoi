using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot("styles", Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = false)]
    public class CT_Styles
    {

        private CT_DocDefaults docDefaultsField;

        private CT_LatentStyles latentStylesField;

        private List<CT_Style> styleField;

        public CT_Styles()
        {
            //this.styleField = new List<CT_Style>();
            //this.latentStylesField = new CT_LatentStyles();
            //this.docDefaultsField = new CT_DocDefaults();
        }
        public static CT_Styles Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Styles ctObj = new CT_Styles();
            ctObj.style = new List<CT_Style>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "docDefaults")
                    ctObj.docDefaults = CT_DocDefaults.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "latentStyles")
                    ctObj.latentStyles = CT_LatentStyles.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "style")
                    ctObj.style.Add(CT_Style.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<w:styles xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
            if (this.docDefaults != null)
                this.docDefaults.Write(sw, "docDefaults");
            if (this.latentStyles != null)
                this.latentStyles.Write(sw, "latentStyles");
            if (this.style != null)
            {
                foreach (CT_Style x in this.style)
                {
                    x.Write(sw, "style");
                }
            }
            sw.Write("</w:styles>");
        }

        [XmlElement(Order = 0)]
        public CT_DocDefaults docDefaults
        {
            get
            {
                return this.docDefaultsField;
            }
            set
            {
                this.docDefaultsField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_LatentStyles latentStyles
        {
            get
            {
                return this.latentStylesField;
            }
            set
            {
                this.latentStylesField = value;
            }
        }

        [XmlElement("style", Order = 2)]
        public List<CT_Style> style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }

        public IList<CT_Style> GetStyleList()
        {
            if(this.styleField==null)
                this.styleField = new List<CT_Style>();
            return style;
        }

        public CT_Style AddNewStyle()
        {
            CT_Style s = new CT_Style();
            if (styleField == null)
                styleField = new List<CT_Style>();
            styleField.Add(s);
            return s;
        }

        public void SetStyleArray(int pos, CT_Style cT_Style)
        {
            lock (this)
            {
                this.styleField[pos] = cT_Style;
            }
        }

        public bool IsSetDocDefaults()
        {
            return this.docDefaultsField != null;
        }

        public CT_DocDefaults AddNewDocDefaults()
        {
            this.docDefaultsField = new CT_DocDefaults();
            return this.docDefaultsField;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DocDefaults
    {

        private CT_RPrDefault rPrDefaultField;

        private CT_PPrDefault pPrDefaultField;

        public CT_DocDefaults()
        {
            //this.pPrDefaultField = new CT_PPrDefault();
            //this.rPrDefaultField = new CT_RPrDefault();
        }
        public static CT_DocDefaults Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DocDefaults ctObj = new CT_DocDefaults();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rPrDefault")
                    ctObj.rPrDefault = CT_RPrDefault.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pPrDefault")
                    ctObj.pPrDefault = CT_PPrDefault.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.rPrDefault != null)
                this.rPrDefault.Write(sw, "rPrDefault");
            if (this.pPrDefault != null)
                this.pPrDefault.Write(sw, "pPrDefault");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_RPrDefault rPrDefault
        {
            get
            {
                return this.rPrDefaultField;
            }
            set
            {
                this.rPrDefaultField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_PPrDefault pPrDefault
        {
            get
            {
                return this.pPrDefaultField;
            }
            set
            {
                this.pPrDefaultField = value;
            }
        }

        public bool IsSetRPrDefault()
        {
            return this.rPrDefaultField != null;
        }

        public CT_RPrDefault AddNewRPrDefault()
        {
            this.rPrDefaultField = new CT_RPrDefault();
            return this.rPrDefaultField;
        }

        public bool IsSetPPrDefault()
        {
            return this.pPrDefaultField != null;
        }

        public CT_PPrDefault AddNewPPrDefault()
        {
            this.pPrDefaultField = new CT_PPrDefault();
            return this.pPrDefaultField;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_RPrDefault
    {

        private CT_RPr rPrField;

        public CT_RPrDefault()
        {
            
        }
        public static CT_RPrDefault Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RPrDefault ctObj = new CT_RPrDefault();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rPr")
                    ctObj.rPr = CT_RPr.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.rPr != null)
                this.rPr.Write(sw, "rPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_RPr rPr
        {
            get
            {
                return this.rPrField;
            }
            set
            {
                this.rPrField = value;
            }
        }

        public bool IsSetRPr()
        {
            return this.rPrField != null;
        }

        public CT_RPr AddNewRPr()
        {
            this.rPrField = new CT_RPr();
            return this.rPrField;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PPrDefault
    {

        private CT_PPr pPrField;

        public CT_PPrDefault()
        {
        }
        public static CT_PPrDefault Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PPrDefault ctObj = new CT_PPrDefault();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "pPr")
                    ctObj.pPr = CT_PPr.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.pPr != null)
                this.pPr.Write(sw, "pPr");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_PPr pPr
        {
            get
            {
                return this.pPrField;
            }
            set
            {
                this.pPrField = value;
            }
        }

        public CT_PPr AddNewPPr()
        {
            this.pPrField = new CT_PPr();
            return this.pPrField;
        }

        public bool IsSetPPr()
        {
            return this.pPrField != null;
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LatentStyles
    {

        private List<CT_LsdException> lsdExceptionField;

        private ST_OnOff defLockedStateField;

        private bool defLockedStateFieldSpecified;

        private string defUIPriorityField;

        private ST_OnOff defSemiHiddenField;

        private bool defSemiHiddenFieldSpecified;

        private ST_OnOff defUnhideWhenUsedField;

        private bool defUnhideWhenUsedFieldSpecified;

        private ST_OnOff defQFormatField;

        private bool defQFormatFieldSpecified;

        private string countField;
        public static CT_LatentStyles Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_LatentStyles ctObj = new CT_LatentStyles();
            if (node.Attributes["w:defLockedState"] != null)
                ctObj.defLockedState = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:defLockedState"].Value);
            ctObj.defUIPriority = XmlHelper.ReadString(node.Attributes["w:defUIPriority"]);
            if (node.Attributes["w:defSemiHidden"] != null)
                ctObj.defSemiHidden = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:defSemiHidden"].Value);
            if (node.Attributes["w:defUnhideWhenUsed"] != null)
                ctObj.defUnhideWhenUsed = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:defUnhideWhenUsed"].Value);
            if (node.Attributes["w:defQFormat"] != null)
                ctObj.defQFormat = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:defQFormat"].Value);
            ctObj.count = XmlHelper.ReadString(node.Attributes["w:count"]);
            ctObj.lsdException = new List<CT_LsdException>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "lsdException")
                    ctObj.lsdException.Add(CT_LsdException.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:defLockedState", this.defLockedState.ToString());
            XmlHelper.WriteAttribute(sw, "w:defUIPriority", this.defUIPriority);
            XmlHelper.WriteAttribute(sw, "w:defSemiHidden", this.defSemiHidden.ToString());
            XmlHelper.WriteAttribute(sw, "w:defUnhideWhenUsed", this.defUnhideWhenUsed.ToString());
            XmlHelper.WriteAttribute(sw, "w:defQFormat", this.defQFormat.ToString());
            XmlHelper.WriteAttribute(sw, "w:count", this.count);
            sw.Write(">");
            if (this.lsdException != null)
            {
                foreach (CT_LsdException x in this.lsdException)
                {
                    x.Write(sw, "lsdException");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        public CT_LatentStyles()
        {
            //this.lsdExceptionField = new List<CT_LsdException>();
        }

        [XmlElement("lsdException", Order = 0)]
        public List<CT_LsdException> lsdException
        {
            get
            {
                return this.lsdExceptionField;
            }
            set
            {
                this.lsdExceptionField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff defLockedState
        {
            get
            {
                return this.defLockedStateField;
            }
            set
            {
                this.defLockedStateField = value;
            }
        }

        [XmlIgnore]
        public bool defLockedStateSpecified
        {
            get
            {
                return this.defLockedStateFieldSpecified;
            }
            set
            {
                this.defLockedStateFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string defUIPriority
        {
            get
            {
                return this.defUIPriorityField;
            }
            set
            {
                this.defUIPriorityField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff defSemiHidden
        {
            get
            {
                return this.defSemiHiddenField;
            }
            set
            {
                this.defSemiHiddenField = value;
            }
        }

        [XmlIgnore]
        public bool defSemiHiddenSpecified
        {
            get
            {
                return this.defSemiHiddenFieldSpecified;
            }
            set
            {
                this.defSemiHiddenFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff defUnhideWhenUsed
        {
            get
            {
                return this.defUnhideWhenUsedField;
            }
            set
            {
                this.defUnhideWhenUsedField = value;
            }
        }

        [XmlIgnore]
        public bool defUnhideWhenUsedSpecified
        {
            get
            {
                return this.defUnhideWhenUsedFieldSpecified;
            }
            set
            {
                this.defUnhideWhenUsedFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff defQFormat
        {
            get
            {
                return this.defQFormatField;
            }
            set
            {
                this.defQFormatField = value;
            }
        }

        [XmlIgnore]
        public bool defQFormatSpecified
        {
            get
            {
                return this.defQFormatFieldSpecified;
            }
            set
            {
                this.defQFormatFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        public CT_LsdException AddNewLsdException()
        {
            CT_LsdException lsd = new CT_LsdException();
            if (this.lsdExceptionField == null)
                this.lsdExceptionField = new List<CT_LsdException>();
            this.lsdExceptionField.Add(lsd);
            return lsd;
        }

        public int SizeOfLsdExceptionArray()
        {
            return lsdExceptionField == null ? 0 : lsdExceptionField.Count;
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Style
    {

        private CT_String nameField;

        private CT_String aliasesField;

        private CT_String basedOnField;

        private CT_String nextField;

        private CT_String linkField;

        private CT_OnOff autoRedefineField;

        private CT_OnOff hiddenField;

        private CT_DecimalNumber uiPriorityField;

        private CT_OnOff semiHiddenField;

        private CT_OnOff unhideWhenUsedField;

        private CT_OnOff qFormatField;

        private CT_OnOff lockedField;

        private CT_OnOff personalField;

        private CT_OnOff personalComposeField;

        private CT_OnOff personalReplyField;

        private CT_LongHexNumber rsidField;

        private CT_PPr pPrField;

        private CT_RPr rPrField;

        private CT_TblPrBase tblPrField;

        private CT_TrPr trPrField;

        private CT_TcPr tcPrField;

        private List<CT_TblStylePr> tblStylePrField;

        private ST_StyleType typeField;

        private bool typeFieldSpecified;

        private string styleIdField;

        private ST_OnOff defaultField;

        private bool defaultFieldSpecified;

        private ST_OnOff customStyleField;

        private bool customStyleFieldSpecified;

        public CT_Style()
        {
            //this.tblStylePrField = new List<CT_TblStylePr>();
            //this.tcPrField = new CT_TcPr();
            //this.trPrField = new CT_TrPr();
            //this.tblPrField = new CT_TblPrBase();
            //this.rPrField = new CT_RPr();
            //this.pPrField = new CT_PPr();
            //this.rsidField = new CT_LongHexNumber();
            //this.personalReplyField = new CT_OnOff();
            //this.personalComposeField = new CT_OnOff();
            //this.personalField = new CT_OnOff();
            //this.lockedField = new CT_OnOff();
            //this.qFormatField = new CT_OnOff();
            //this.unhideWhenUsedField = new CT_OnOff();
            //this.semiHiddenField = new CT_OnOff();
            //this.uiPriorityField = new CT_DecimalNumber();
            //this.hiddenField = new CT_OnOff();
            //this.autoRedefineField = new CT_OnOff();
            //this.linkField = new CT_String();
            //this.nextField = new CT_String();
            //this.basedOnField = new CT_String();
            //this.aliasesField = new CT_String();
            //this.nameField = new CT_String();
        }
        public static CT_Style Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Style ctObj = new CT_Style();
            if (node.Attributes["w:type"] != null)
                ctObj.type = (ST_StyleType)Enum.Parse(typeof(ST_StyleType), node.Attributes["w:type"].Value);
            ctObj.styleId = XmlHelper.ReadString(node.Attributes["w:styleId"]);
            if (node.Attributes["w:default"] != null)
                ctObj.@default = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:default"].Value);
            if (node.Attributes["w:customStyle"] != null)
                ctObj.customStyle = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:customStyle"].Value);
            ctObj.tblStylePr = new List<CT_TblStylePr>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "name")
                    ctObj.name = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "aliases")
                    ctObj.aliases = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "basedOn")
                    ctObj.basedOn = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "next")
                    ctObj.next = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "link")
                    ctObj.link = CT_String.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "autoRedefine")
                    ctObj.autoRedefine = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "hidden")
                    ctObj.hidden = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "uiPriority")
                    ctObj.uiPriority = CT_DecimalNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "semiHidden")
                    ctObj.semiHidden = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "unhideWhenUsed")
                    ctObj.unhideWhenUsed = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "qFormat")
                    ctObj.qFormat = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "locked")
                    ctObj.locked = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "personal")
                    ctObj.personal = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "personalCompose")
                    ctObj.personalCompose = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "personalReply")
                    ctObj.personalReply = CT_OnOff.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rsid")
                    ctObj.rsid = CT_LongHexNumber.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "pPr")
                    ctObj.pPr = CT_PPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "rPr")
                    ctObj.rPr = CT_RPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblPr")
                    ctObj.tblPr = CT_TblPrBase.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "trPr")
                    ctObj.trPr = CT_TrPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tcPr")
                    ctObj.tcPr = CT_TcPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tblStylePr")
                    ctObj.tblStylePr.Add(CT_TblStylePr.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:type", this.type.ToString());
            if(this.@default!= ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:default", this.@default.ToString());
            if (this.customStyle == ST_OnOff.on)
                XmlHelper.WriteAttribute(sw, "w:customStyle", "1");
            XmlHelper.WriteAttribute(sw, "w:styleId", this.styleId);
            sw.Write(">");
            if (this.name != null)
                this.name.Write(sw, "name");
            if (this.aliases != null)
                this.aliases.Write(sw, "aliases");
            if (this.basedOn != null)
                this.basedOn.Write(sw, "basedOn");
            if (this.next != null)
                this.next.Write(sw, "next");
            if (this.link != null)
                this.link.Write(sw, "link");
            if (this.autoRedefine != null)
                this.autoRedefine.Write(sw, "autoRedefine");
            if (this.hidden != null)
                this.hidden.Write(sw, "hidden");
            if (this.uiPriority != null)
                this.uiPriority.Write(sw, "uiPriority");
            if (this.semiHidden != null)
                this.semiHidden.Write(sw, "semiHidden");
            if (this.unhideWhenUsed != null)
                this.unhideWhenUsed.Write(sw, "unhideWhenUsed");
            if (this.qFormat != null)
                this.qFormat.Write(sw, "qFormat");
            if (this.locked != null)
                this.locked.Write(sw, "locked");
            if (this.personal != null)
                this.personal.Write(sw, "personal");
            if (this.personalCompose != null)
                this.personalCompose.Write(sw, "personalCompose");
            if (this.personalReply != null)
                this.personalReply.Write(sw, "personalReply");
            if (this.rsid != null)
                this.rsid.Write(sw, "rsid");
            if (this.pPr != null)
                this.pPr.Write(sw, "pPr");
            if (this.rPr != null)
                this.rPr.Write(sw, "rPr");
            if (this.tblPr != null)
                this.tblPr.Write(sw, "tblPr");
            if (this.trPr != null)
                this.trPr.Write(sw, "trPr");
            if (this.tcPr != null)
                this.tcPr.Write(sw, "tcPr");
            if (this.tblStylePr != null)
            {
                foreach (CT_TblStylePr x in this.tblStylePr)
                {
                    x.Write(sw, "tblStylePr");
                }
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlElement(Order = 0)]
        public CT_String name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_String aliases
        {
            get
            {
                return this.aliasesField;
            }
            set
            {
                this.aliasesField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_String basedOn
        {
            get
            {
                return this.basedOnField;
            }
            set
            {
                this.basedOnField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_String next
        {
            get
            {
                return this.nextField;
            }
            set
            {
                this.nextField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_String link
        {
            get
            {
                return this.linkField;
            }
            set
            {
                this.linkField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_OnOff autoRedefine
        {
            get
            {
                return this.autoRedefineField;
            }
            set
            {
                this.autoRedefineField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_OnOff hidden
        {
            get
            {
                return this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_DecimalNumber uiPriority
        {
            get
            {
                return this.uiPriorityField;
            }
            set
            {
                this.uiPriorityField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_OnOff semiHidden
        {
            get
            {
                return this.semiHiddenField;
            }
            set
            {
                this.semiHiddenField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_OnOff unhideWhenUsed
        {
            get
            {
                return this.unhideWhenUsedField;
            }
            set
            {
                this.unhideWhenUsedField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_OnOff qFormat
        {
            get
            {
                return this.qFormatField;
            }
            set
            {
                this.qFormatField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_OnOff locked
        {
            get
            {
                return this.lockedField;
            }
            set
            {
                this.lockedField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_OnOff personal
        {
            get
            {
                return this.personalField;
            }
            set
            {
                this.personalField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_OnOff personalCompose
        {
            get
            {
                return this.personalComposeField;
            }
            set
            {
                this.personalComposeField = value;
            }
        }

        [XmlElement(Order = 14)]
        public CT_OnOff personalReply
        {
            get
            {
                return this.personalReplyField;
            }
            set
            {
                this.personalReplyField = value;
            }
        }

        [XmlElement(Order = 15)]
        public CT_LongHexNumber rsid
        {
            get
            {
                return this.rsidField;
            }
            set
            {
                this.rsidField = value;
            }
        }

        [XmlElement(Order = 16)]
        public CT_PPr pPr
        {
            get
            {
                return this.pPrField;
            }
            set
            {
                this.pPrField = value;
            }
        }

        [XmlElement(Order = 17)]
        public CT_RPr rPr
        {
            get
            {
                return this.rPrField;
            }
            set
            {
                this.rPrField = value;
            }
        }

        [XmlElement(Order = 18)]
        public CT_TblPrBase tblPr
        {
            get
            {
                return this.tblPrField;
            }
            set
            {
                this.tblPrField = value;
            }
        }

        [XmlElement(Order = 19)]
        public CT_TrPr trPr
        {
            get
            {
                return this.trPrField;
            }
            set
            {
                this.trPrField = value;
            }
        }

        [XmlElement(Order = 20)]
        public CT_TcPr tcPr
        {
            get
            {
                return this.tcPrField;
            }
            set
            {
                this.tcPrField = value;
            }
        }

        [XmlElement("tblStylePr", Order = 21)]
        public List<CT_TblStylePr> tblStylePr
        {
            get
            {
                return this.tblStylePrField;
            }
            set
            {
                this.tblStylePrField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_StyleType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [XmlIgnore]
        public bool typeSpecified
        {
            get
            {
                return this.typeFieldSpecified;
            }
            set
            {
                this.typeFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string styleId
        {
            get
            {
                return this.styleIdField;
            }
            set
            {
                this.styleIdField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff @default
        {
            get
            {
                return this.defaultField;
            }
            set
            {
                this.defaultField = value;
            }
        }

        [XmlIgnore]
        public bool defaultSpecified
        {
            get
            {
                return this.defaultFieldSpecified;
            }
            set
            {
                this.defaultFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff customStyle
        {
            get
            {
                return this.customStyleField;
            }
            set
            {
                this.customStyleField = value;
            }
        }

        [XmlIgnore]
        public bool customStyleSpecified
        {
            get
            {
                return this.customStyleFieldSpecified;
            }
            set
            {
                this.customStyleFieldSpecified = value;
            }
        }

        public bool IsSetName()
        {
            return this.name != null;
        }
    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Em
    {
        public static CT_Em Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Em ctObj = new CT_Em();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_Em)Enum.Parse(typeof(ST_Em), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        private ST_Em valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Em val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_VerticalJc
    {
        public static CT_VerticalJc Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_VerticalJc ctObj = new CT_VerticalJc();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_VerticalJc)Enum.Parse(typeof(ST_VerticalJc), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write("/>");
        }


        private ST_VerticalJc valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_VerticalJc val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_VerticalJc
    {


        top,


        center,


        both,


        bottom,
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Em
    {


        none,


        dot,


        comma,


        circle,


        underDot,
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Shd
    {

        private ST_Shd valField;

        private string colorField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;

        private string fillField;

        private ST_ThemeColor themeFillField;

        private bool themeFillFieldSpecified;

        private byte[] themeFillTintField;

        private byte[] themeFillShadeField;
        public CT_Shd()
        {
            this.themeColorField = ST_ThemeColor.none;
            this.themeFillField = ST_ThemeColor.none;
        }

        public static CT_Shd Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Shd ctObj = new CT_Shd();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_Shd)Enum.Parse(typeof(ST_Shd), node.Attributes["w:val"].Value);
            ctObj.color = XmlHelper.ReadString(node.Attributes["w:color"]);
            if (node.Attributes["w:themeColor"] != null)
                ctObj.themeColor = (ST_ThemeColor)Enum.Parse(typeof(ST_ThemeColor), node.Attributes["w:themeColor"].Value);
            ctObj.themeTint = XmlHelper.ReadBytes(node.Attributes["w:themeTint"]);
            ctObj.themeShade = XmlHelper.ReadBytes(node.Attributes["w:themeShade"]);
            ctObj.fill = XmlHelper.ReadString(node.Attributes["w:fill"]);
            if (node.Attributes["w:themeFill"] != null)
                ctObj.themeFill = (ST_ThemeColor)Enum.Parse(typeof(ST_ThemeColor), node.Attributes["w:themeFill"].Value);
            ctObj.themeFillTint = XmlHelper.ReadBytes(node.Attributes["w:themeFillTint"]);
            ctObj.themeFillShade = XmlHelper.ReadBytes(node.Attributes["w:themeFillShade"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            XmlHelper.WriteAttribute(sw, "w:color", this.color);
            if(this.themeColor!= ST_ThemeColor.none)
                XmlHelper.WriteAttribute(sw, "w:themeColor", this.themeColor.ToString());
            XmlHelper.WriteAttribute(sw, "w:themeTint", this.themeTint);
            XmlHelper.WriteAttribute(sw, "w:themeShade", this.themeShade);
            XmlHelper.WriteAttribute(sw, "w:fill", this.fill);
            if(this.themeFill!= ST_ThemeColor.none)
                XmlHelper.WriteAttribute(sw, "w:themeFill", this.themeFill.ToString());
            XmlHelper.WriteAttribute(sw, "w:themeFillTint", this.themeFillTint);
            XmlHelper.WriteAttribute(sw, "w:themeFillShade", this.themeFillShade);
            sw.Write("/>");
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Shd val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string color
        {
            get
            {
                return this.colorField;
            }
            set
            {
                this.colorField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeColor
        {
            get
            {
                return this.themeColorField;
            }
            set
            {
                this.themeColorField = value;
            }
        }

        [XmlIgnore]
        public bool themeColorSpecified
        {
            get
            {
                return this.themeColorFieldSpecified;
            }
            set
            {
                this.themeColorFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeTint
        {
            get
            {
                return this.themeTintField;
            }
            set
            {
                this.themeTintField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeShade
        {
            get
            {
                return this.themeShadeField;
            }
            set
            {
                this.themeShadeField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string fill
        {
            get
            {
                return this.fillField;
            }
            set
            {
                this.fillField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeFill
        {
            get
            {
                return this.themeFillField;
            }
            set
            {
                this.themeFillField = value;
            }
        }

        [XmlIgnore]
        public bool themeFillSpecified
        {
            get
            {
                return this.themeFillFieldSpecified;
            }
            set
            {
                this.themeFillFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeFillTint
        {
            get
            {
                return this.themeFillTintField;
            }
            set
            {
                this.themeFillTintField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeFillShade
        {
            get
            {
                return this.themeFillShadeField;
            }
            set
            {
                this.themeFillShadeField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Shd
    {

    
        nil,

    
        clear,

    
        solid,

    
        horzStripe,

    
        vertStripe,

    
        reverseDiagStripe,

    
        diagStripe,

    
        horzCross,

    
        diagCross,

    
        thinHorzStripe,

    
        thinVertStripe,

    
        thinReverseDiagStripe,

    
        thinDiagStripe,

    
        thinHorzCross,

    
        thinDiagCross,

    
        pct5,

    
        pct10,

    
        pct12,

    
        pct15,

    
        pct20,

    
        pct25,

    
        pct30,

    
        pct35,

    
        pct37,

    
        pct40,

    
        pct45,

    
        pct50,

    
        pct55,

    
        pct60,

    
        pct62,

    
        pct65,

    
        pct70,

    
        pct75,

    
        pct80,

    
        pct85,

    
        pct87,

    
        pct90,

    
        pct95,
    }










    //==============


    /// <summary>
    /// Text Expansion/Compression Percentage
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextScale
    {
        public static CT_TextScale Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TextScale ctObj = new CT_TextScale();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        private string valField;
        /// <summary>
        /// Text Expansion/Compression Value
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    /// <summary>
    /// Text Highlight Colors
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Highlight
    {
        public static CT_Highlight Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Highlight ctObj = new CT_Highlight();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_HighlightColor)Enum.Parse(typeof(ST_HighlightColor), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        private ST_HighlightColor valField;
        /// <summary>
        /// Highlighting Color
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_HighlightColor val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }
    /// <summary>
    /// Color Value
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Color
    {

        private string valField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;

        public CT_Color()
        {
            this.themeColorField = ST_ThemeColor.none;
        }

        public static CT_Color Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Color ctObj = new CT_Color();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            if (node.Attributes["w:themeColor"] != null)
                ctObj.themeColor = (ST_ThemeColor)Enum.Parse(typeof(ST_ThemeColor), node.Attributes["w:themeColor"].Value);
            ctObj.themeTint = XmlHelper.ReadBytes(node.Attributes["w:themeTint"]);
            ctObj.themeShade = XmlHelper.ReadBytes(node.Attributes["w:themeShade"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val, true);
            if(this.themeColorField!= ST_ThemeColor.none)
                XmlHelper.WriteAttribute(sw, "w:themeColor", this.themeColor.ToString());
            XmlHelper.WriteAttribute(sw, "w:themeTint", this.themeTint);
            XmlHelper.WriteAttribute(sw, "w:themeShade", this.themeShade);
            sw.Write("/>");
        }

        /// <summary>
        /// Run Content Color
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
        /// <summary>
        /// Run Content Theme Color
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeColor
        {
            get
            {
                return this.themeColorField;
            }
            set
            {
                this.themeColorField = value;
            }
        }

        [XmlIgnore]
        public bool themeColorSpecified
        {
            get
            {
                return this.themeColorFieldSpecified;
            }
            set
            {
                this.themeColorFieldSpecified = value;
            }
        }
        /// <summary>
        /// Run Content Theme Color Tint
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeTint
        {
            get
            {
                return this.themeTintField;
            }
            set
            {
                this.themeTintField = value;
            }
        }
        /// <summary>
        /// Run Content Theme Color Shade
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeShade
        {
            get
            {
                return this.themeShadeField;
            }
            set
            {
                this.themeShadeField = value;
            }
        }
    }

    /// <summary>
    /// Underline Style
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Underline
    {

        private ST_Underline valField;

        private bool valFieldSpecified;

        private string colorField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;

        public CT_Underline()
        {
            this.themeColorField = ST_ThemeColor.none;
        }

        public static CT_Underline Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Underline ctObj = new CT_Underline();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_Underline)Enum.Parse(typeof(ST_Underline), node.Attributes["w:val"].Value);
            ctObj.color = XmlHelper.ReadString(node.Attributes["w:color"]);
            if (node.Attributes["w:themeColor"] != null)
                ctObj.themeColor = (ST_ThemeColor)Enum.Parse(typeof(ST_ThemeColor), node.Attributes["w:themeColor"].Value);
            ctObj.themeTint = XmlHelper.ReadBytes(node.Attributes["w:themeTint"]);
            ctObj.themeShade = XmlHelper.ReadBytes(node.Attributes["w:themeShade"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            XmlHelper.WriteAttribute(sw, "w:color", this.color);
            if(this.themeColor!= ST_ThemeColor.none)
                XmlHelper.WriteAttribute(sw, "w:themeColor", this.themeColor.ToString());
            XmlHelper.WriteAttribute(sw, "w:themeTint", this.themeTint);
            XmlHelper.WriteAttribute(sw, "w:themeShade", this.themeShade);
            sw.Write("/>");
        }


        /// <summary>
        /// Underline Style value
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Underline val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }

        [XmlIgnore]
        public bool valSpecified
        {
            get
            {
                return this.valFieldSpecified;
            }
            set
            {
                this.valFieldSpecified = value;
            }
        }
        /// <summary>
        /// Underline Color
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string color
        {
            get
            {
                return this.colorField;
            }
            set
            {
                this.colorField = value;
            }
        }
        /// <summary>
        /// Underline Theme Color
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeColor
        {
            get
            {
                return this.themeColorField;
            }
            set
            {
                this.themeColorField = value;
            }
        }

        [XmlIgnore]
        public bool themeColorSpecified
        {
            get
            {
                return this.themeColorFieldSpecified;
            }
            set
            {
                this.themeColorFieldSpecified = value;
            }
        }
        /// <summary>
        /// Underline Theme Color Tint
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeTint
        {
            get
            {
                return this.themeTintField;
            }
            set
            {
                this.themeTintField = value;
            }
        }
        /// <summary>
        /// Underline Theme Color Shade
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeShade
        {
            get
            {
                return this.themeShadeField;
            }
            set
            {
                this.themeShadeField = value;
            }
        }
    }

    /// <summary>
    /// Underline Patterns
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Underline
    {

        /// <summary>
        /// Single Underline
        /// </summary>
        single,

        /// <summary>
        /// Underline Non-Space Characters Only
        /// </summary>
        words,

        /// <summary>
        /// Double Underline
        /// </summary>
        @double,

        /// <summary>
        /// Thick Underline
        /// </summary>
        thick,

        /// <summary>
        /// Dotted Underline
        /// </summary>
        dotted,

        /// <summary>
        /// Thick Dotted Underline
        /// </summary>
        dottedHeavy,

        /// <summary>
        /// Dashed Underline
        /// </summary>
        dash,

        /// <summary>
        /// Thick Dashed Underline
        /// </summary>
        dashedHeavy,

        /// <summary>
        /// Long Dashed Underline
        /// </summary>
        dashLong,

        /// <summary>
        /// Thick Long Dashed Underline
        /// </summary>
        dashLongHeavy,

        /// <summary>
        /// Dash-Dot Underline
        /// </summary>
        dotDash,

        /// <summary>
        /// Thick Dash-Dot Underline
        /// </summary>
        dashDotHeavy,

        /// <summary>
        /// Dash-Dot-Dot Underline
        /// </summary>
        dotDotDash,

        /// <summary>
        /// Thick Dash-Dot-Dot Underline
        /// </summary>
        dashDotDotHeavy,

        /// <summary>
        /// Wave Underline
        /// </summary>
        wave,

        /// <summary>
        /// Heavy Wave Underline
        /// </summary>
        wavyHeavy,

        /// <summary>
        /// Double Wave Underline
        /// </summary>
        wavyDouble,

        /// <summary>
        /// No Underline
        /// </summary>
        none,
    }

    /// <summary>
    /// Animated Text Effects
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TextEffect
    {
        public static CT_TextEffect Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TextEffect ctObj = new CT_TextEffect();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_TextEffect)Enum.Parse(typeof(ST_TextEffect), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write("/>");
        }


        private ST_TextEffect valField;
        /// <summary>
        /// Animated Text Effect Type
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TextEffect val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TextEffect
    {

        /// <summary>
        /// Blinking Background Animation
        /// </summary>
        blinkBackground,

        /// <summary>
        /// Colored Lights Animation
        /// </summary>
        lights,

        /// <summary>
        /// Black Dashed Line Animation
        /// </summary>
        antsBlack,

        /// <summary>
        /// Marching Red Ants
        /// </summary>
        antsRed,

        /// <summary>
        /// Shimmer Animation
        /// </summary>
        shimmer,

        /// <summary>
        /// Sparkling Lights Animation
        /// </summary>
        sparkle,

        /// <summary>
        /// No Animation
        /// </summary>
        none,
    }

    /// <summary>
    /// Border Style
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Border
    {

        private ST_Border valField;

        private string colorField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;

        private ulong szField;

        private bool szFieldSpecified;

        private ulong spaceField;

        private bool spaceFieldSpecified;

        private ST_OnOff shadowField;

        private bool shadowFieldSpecified;

        private ST_OnOff frameField;

        private bool frameFieldSpecified;

        public CT_Border()
        {
            this.themeColor = ST_ThemeColor.none;
        }

        public static CT_Border Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Border ctObj = new CT_Border();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_Border)Enum.Parse(typeof(ST_Border), node.Attributes["w:val"].Value);
            ctObj.color = XmlHelper.ReadString(node.Attributes["w:color"]);
            if (node.Attributes["w:themeColor"] != null)
                ctObj.themeColor = (ST_ThemeColor)Enum.Parse(typeof(ST_ThemeColor), node.Attributes["w:themeColor"].Value);
            ctObj.themeTint = XmlHelper.ReadBytes(node.Attributes["w:themeTint"]);
            ctObj.themeShade = XmlHelper.ReadBytes(node.Attributes["w:themeShade"]);
            ctObj.sz = XmlHelper.ReadULong(node.Attributes["w:sz"]);
            ctObj.space = XmlHelper.ReadULong(node.Attributes["w:space"]);
            if (node.Attributes["w:shadow"] != null)
                ctObj.shadow = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:shadow"].Value);
            if (node.Attributes["w:frame"] != null)
                ctObj.frame = (ST_OnOff)Enum.Parse(typeof(ST_OnOff), node.Attributes["w:frame"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            if (this.themeColor != ST_ThemeColor.none)
                XmlHelper.WriteAttribute(sw, "w:themeColor", this.themeColor.ToString());
            XmlHelper.WriteAttribute(sw, "w:themeTint", this.themeTint);
            XmlHelper.WriteAttribute(sw, "w:themeShade", this.themeShade);
            XmlHelper.WriteAttribute(sw, "w:sz", this.sz);
            XmlHelper.WriteAttribute(sw, "w:space", this.space, true);
            XmlHelper.WriteAttribute(sw, "w:color", this.color);
            if(this.shadow!= ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:shadow", this.shadow.ToString());
            if (this.frame != ST_OnOff.off)
                XmlHelper.WriteAttribute(sw, "w:frame", this.frame.ToString());
            sw.Write("/>");
        }



        /// <summary>
        /// Border Style
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Border val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
        /// <summary>
        /// Border Color
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string color
        {
            get
            {
                return this.colorField;
            }
            set
            {
                this.colorField = value;
            }
        }
        /// <summary>
        /// Border Theme Color
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeColor
        {
            get
            {
                return this.themeColorField;
            }
            set
            {
                this.themeColorField = value;
            }
        }

        [XmlIgnore]
        public bool themeColorSpecified
        {
            get
            {
                return this.themeColorFieldSpecified;
            }
            set
            {
                this.themeColorFieldSpecified = value;
            }
        }
        /// <summary>
        /// Border Theme Color Tint
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeTint
        {
            get
            {
                return this.themeTintField;
            }
            set
            {
                this.themeTintField = value;
            }
        }
        /// <summary>
        /// Border Theme Color Shade
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeShade
        {
            get
            {
                return this.themeShadeField;
            }
            set
            {
                this.themeShadeField = value;
            }
        }
        /// <summary>
        /// Border Width
        /// </summary>
        /// ST_EighthPointMeasure
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong sz
        {
            get
            {
                return this.szField;
            }
            set
            {
                this.szField = value;
            }
        }

        [XmlIgnore]
        public bool szSpecified
        {
            get
            {
                return this.szFieldSpecified;
            }
            set
            {
                this.szFieldSpecified = value;
            }
        }
        /// <summary>
        /// Border Spacing Measurement
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong space
        {
            get
            {
                return this.spaceField;
            }
            set
            {
                this.spaceField = value;
            }
        }

        [XmlIgnore]
        public bool spaceSpecified
        {
            get
            {
                return this.spaceFieldSpecified;
            }
            set
            {
                this.spaceFieldSpecified = value;
            }
        }
        /// <summary>
        /// Border Shadow
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff shadow
        {
            get
            {
                return this.shadowField;
            }
            set
            {
                this.shadowField = value;
            }
        }

        [XmlIgnore]
        public bool shadowSpecified
        {
            get
            {
                return this.shadowFieldSpecified;
            }
            set
            {
                this.shadowFieldSpecified = value;
            }
        }
        /// <summary>
        /// Create Frame Effect
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff frame
        {
            get
            {
                return this.frameField;
            }
            set
            {
                this.frameField = value;
            }
        }

        [XmlIgnore]
        public bool frameSpecified
        {
            get
            {
                return this.frameFieldSpecified;
            }
            set
            {
                this.frameFieldSpecified = value;
            }
        }
    }

    /// <summary>
    /// Border Styles
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Border
    {

        /// <summary>
        /// No Border
        /// </summary>
        nil,

        /// <summary>
        /// No Border
        /// </summary>
        none,

        /// <summary>
        /// Single Line Border
        /// </summary>
        single,

        /// <summary>
        /// Single Line Border
        /// </summary>
        thick,

        /// <summary>
        /// Double Line Border
        /// </summary>
        @double,

        /// <summary>
        /// Dotted Line Border
        /// </summary>
        dotted,

        /// <summary>
        /// Dashed Line Border
        /// </summary>
        dashed,

        /// <summary>
        /// Dot Dash Line Border
        /// </summary>
        dotDash,

        /// <summary>
        /// Dot Dot Dash Line Border
        /// </summary>
        dotDotDash,

        /// <summary>
        /// Triple Line Border
        /// </summary>
        triple,

        /// <summary>
        /// Thin, Thick Line Border<
        /// </summary>
        thinThickSmallGap,

        /// <summary>
        /// Thick, Thin Line Border
        /// </summary>
        thickThinSmallGap,

        /// <summary>
        /// Thin, Thick, Thin Line Border
        /// </summary>
        thinThickThinSmallGap,

        /// <summary>
        /// Thin, Thick Line Border
        /// </summary>
        thinThickMediumGap,

        /// <summary>
        /// Thick, Thin Line Border
        /// </summary>
        thickThinMediumGap,

        /// <summary>
        /// Thin, Thick, Thin Line Border
        /// </summary>
        thinThickThinMediumGap,

        /// <summary>
        /// Thin, Thick Line Border
        /// </summary>
        thinThickLargeGap,

        /// <summary>
        /// Thick, Thin Line Border
        /// </summary>
        thickThinLargeGap,

        /// <summary>
        /// Thin, Thick, Thin Line Border
        /// </summary>
        thinThickThinLargeGap,

        /// <summary>
        /// Wavy Line Border
        /// </summary>
        wave,

        /// <summary>
        /// Double Wave Line Border
        /// </summary>
        doubleWave,

        /// <summary>
        /// Dashed Line Border
        /// </summary>
        dashSmallGap,

        /// <summary>
        /// Dash Dot Strokes Line Border
        /// </summary>
        dashDotStroked,

        /// <summary>
        /// 3D Embossed Line Border
        /// </summary>
        threeDEmboss,

        /// <summary>
        /// 3D Engraved Line Border
        /// </summary>
        threeDEngrave,

        /// <summary>
        /// Outset Line Border
        /// </summary>
        outset,

        /// <summary>
        /// Inset Line Border
        /// </summary>
        inset,

        /// <summary>
        /// Apples Art Border
        /// </summary>
        apples,

        /// <summary>
        /// Arched Scallops Art Border
        /// </summary>
        archedScallops,

        /// <summary>
        /// Baby Pacifier Art Border
        /// </summary>
        babyPacifier,

        /// <summary>
        /// Baby Rattle Art Border
        /// </summary>
        babyRattle,

        /// <summary>
        /// Three Color Balloons Art Border
        /// </summary>
        balloons3Colors,

        /// <summary>
        /// Hot Air Balloons Art Border
        /// </summary>
        balloonsHotAir,

        /// <summary>
        /// Black Dash Art Border
        /// </summary>
        basicBlackDashes,

        /// <summary>
        /// Black Dot Art Border
        /// </summary>
        basicBlackDots,

        /// <summary>
        /// Black Square Art Border
        /// </summary>
        basicBlackSquares,

        /// <summary>
        /// Thin Line Art Border
        /// </summary>
        basicThinLines,

        /// <summary>
        /// White Dash Art Border
        /// </summary>
        basicWhiteDashes,

        /// <summary>
        /// White Dot Art Border
        /// </summary>
        basicWhiteDots,

        /// <summary>
        /// White Square Art Border
        /// </summary>
        basicWhiteSquares,

        /// <summary>
        /// Wide Inline Art Border
        /// </summary>
        basicWideInline,

        /// <summary>
        /// Wide Midline Art Border
        /// </summary>
        basicWideMidline,

        /// <summary>
        /// 
        /// </summary>
        basicWideOutline,

        /// <summary>
        /// Wide Outline Art Border
        /// </summary>
        bats,

        /// <summary>
        /// Bats Art Border
        /// </summary>
        birds,

        /// <summary>
        /// Birds Art Border
        /// </summary>
        birdsFlight,

        /// <summary>
        /// Cabin Art Border
        /// </summary>
        cabins,

        /// <summary>
        /// Cake Art Border
        /// </summary>
        cakeSlice,

        /// <summary>
        /// Candy Corn Art Border
        /// </summary>
        candyCorn,

        /// <summary>
        /// Knot Work Art Border
        /// </summary>
        celticKnotwork,

        /// <summary>
        /// Certificate Banner Art Border
        /// </summary>
        certificateBanner,

        /// <summary>
        /// Chain Link Art Border
        /// </summary>
        chainLink,

        /// <summary>
        /// Champagne Bottle Art Border
        /// </summary>
        champagneBottle,

        /// <summary>
        /// Black and White Bar Art Border
        /// </summary>
        checkedBarBlack,

        /// <summary>
        /// Color Checked Bar Art Border
        /// </summary>
        checkedBarColor,

        /// <summary>
        /// Checkerboard Art Border
        /// </summary>
        checkered,

        /// <summary>
        /// Christmas Tree Art Border
        /// </summary>
        christmasTree,

        /// <summary>
        /// Circles And Lines Art Border
        /// </summary>
        circlesLines,

        /// <summary>
        /// Circles and Rectangles Art Border
        /// </summary>
        circlesRectangles,

        /// <summary>
        /// Wave Art Border
        /// </summary>
        classicalWave,

        /// <summary>
        /// Clocks Art Border
        /// </summary>
        clocks,

        /// <summary>
        /// Compass Art Border
        /// </summary>
        compass,

        /// <summary>
        /// Confetti Art Border
        /// </summary>
        confetti,

        /// <summary>
        /// Confetti Art Border
        /// </summary>
        confettiGrays,

        /// <summary>
        /// Confetti Art Border
        /// </summary>
        confettiOutline,

        /// <summary>
        /// Confetti Streamers Art Border
        /// </summary>
        confettiStreamers,

        /// <summary>
        /// Confetti Art Border
        /// </summary>
        confettiWhite,

        /// <summary>
        /// Corner Triangle Art Border
        /// </summary>
        cornerTriangles,

        /// <summary>
        /// Dashed Line Art Border
        /// </summary>
        couponCutoutDashes,

        /// <summary>
        /// Dotted Line Art Border
        /// </summary>
        couponCutoutDots,

        /// <summary>
        /// Maze Art Border
        /// </summary>
        crazyMaze,

        /// <summary>
        /// Butterfly Art Border
        /// </summary>
        creaturesButterfly,

        /// <summary>
        /// Fish Art Border
        /// </summary>
        creaturesFish,

        /// <summary>
        /// Insects Art Border
        /// </summary>
        creaturesInsects,

        /// <summary>
        /// Ladybug Art Border
        /// </summary>
        creaturesLadyBug,

        /// <summary>
        /// Cross-stitch Art Border
        /// </summary>
        crossStitch,

        /// <summary>
        /// Cupid Art Border
        /// </summary>
        cup,

        /// <summary>
        /// Archway Art Border
        /// </summary>
        decoArch,

        /// <summary>
        /// Color Archway Art Border
        /// </summary>
        decoArchColor,

        /// <summary>
        /// Blocks Art Border
        /// </summary>
        decoBlocks,

        /// <summary>
        /// Gray Diamond Art Border
        /// </summary>
        diamondsGray,

        /// <summary>
        /// Double D Art Border
        /// </summary>
        doubleD,

        /// <summary>
        /// Diamond Art Border
        /// </summary>
        doubleDiamonds,

        /// <summary>
        /// Earth Art Border
        /// </summary>
        earth1,

        /// <summary>
        /// Earth Art Border
        /// </summary>
        earth2,

        /// <summary>
        /// Shadowed Square Art Border
        /// </summary>
        eclipsingSquares1,

        /// <summary>
        /// Shadowed Square Art Border
        /// </summary>
        eclipsingSquares2,

        /// <summary>
        /// Painted Egg Art Border
        /// </summary>
        eggsBlack,

        /// <summary>
        /// Fans Art Border
        /// </summary>
        fans,

        /// <summary>
        /// Film Reel Art Border
        /// </summary>
        film,

        /// <summary>
        /// Firecracker Art Border
        /// </summary>
        firecrackers,

        /// <summary>
        /// Flowers Art Border
        /// </summary>
        flowersBlockPrint,

        /// <summary>
        /// Daisy Art Border
        /// </summary>
        flowersDaisies,

        /// <summary>
        /// Flowers Art Border
        /// </summary>
        flowersModern1,

        /// <summary>
        /// Flowers Art Border
        /// </summary>
        flowersModern2,

        /// <summary>
        /// Pansy Art Border
        /// </summary>
        flowersPansy,

        /// <summary>
        /// Red Rose Art Border
        /// </summary>
        flowersRedRose,

        /// <summary>
        /// Roses Art Border
        /// </summary>
        flowersRoses,

        /// <summary>
        /// Flowers in a Teacup Art Border
        /// </summary>
        flowersTeacup,

        /// <summary>
        /// Small Flower Art Border
        /// </summary>
        flowersTiny,

        /// <summary>
        /// Gems Art Border
        /// </summary>
        gems,

        /// <summary>
        /// Gingerbread Man Art Border
        /// </summary>
        gingerbreadMan,

        /// <summary>
        /// Triangle Gradient Art Border
        /// </summary>
        gradient,

        /// <summary>
        /// Handmade Art Border
        /// </summary>
        handmade1,

        /// <summary>
        /// Handmade Art Border
        /// </summary>
        handmade2,

        /// <summary>
        /// Heart-Shaped Balloon Art Border
        /// </summary>
        heartBalloon,

        /// <summary>
        /// Gray Heart Art Border
        /// </summary>
        heartGray,

        /// <summary>
        /// Hearts Art Border
        /// </summary>
        hearts,

        /// <summary>
        /// Pattern Art Border
        /// </summary>
        heebieJeebies,

        /// <summary>
        /// Holly Art Border
        /// </summary>
        holly,

        /// <summary>
        /// House Art Border
        /// </summary>
        houseFunky,

        /// <summary>
        /// Circular Art Border
        /// </summary>
        hypnotic,

        /// <summary>
        /// Ice Cream Cone Art Border
        /// </summary>
        iceCreamCones,

        /// <summary>
        /// Light Bulb Art Border
        /// </summary>
        lightBulb,

        /// <summary>
        /// Light Bulb Art Border
        /// </summary>
        lightning1,

        /// <summary>
        /// Light Bulb Art Border
        /// </summary>
        lightning2,

        /// <summary>
        /// Map Pins Art Border
        /// </summary>
        mapPins,

        /// <summary>
        /// Maple Leaf Art Border
        /// </summary>
        mapleLeaf,
        
        /// <summary>
        /// Muffin Art Border
        /// </summary>
        mapleMuffins,

        /// <summary>
        /// Marquee Art Border
        /// </summary>
        marquee,

        /// <summary>
        /// Marquee Art Border
        /// </summary>
        marqueeToothed,

        /// <summary>
        /// Moon Art Border
        /// </summary>
        moons,

        /// <summary>
        /// Mosaic Art Border
        /// </summary>
        mosaic,

        /// <summary>
        /// Musical Note Art Border
        /// </summary>
        musicNotes,

        /// <summary>
        /// Patterned Art Border
        /// </summary>
        northwest,

        /// <summary>
        /// Oval Art Border
        /// </summary>
        ovals,

        /// <summary>
        /// Package Art Border
        /// </summary>
        packages,

        /// <summary>
        /// Black Palm Tree Art Border
        /// </summary>
        palmsBlack,

        /// <summary>
        /// Color Palm Tree Art Border
        /// </summary>
        palmsColor,

        /// <summary>
        /// Paper Clip Art Border
        /// </summary>
        paperClips,

        /// <summary>
        /// Papyrus Art Border
        /// </summary>
        papyrus,

        /// <summary>
        /// Party Favor Art Border
        /// </summary>
        partyFavor,

        /// <summary>
        /// Party Glass Art Border
        /// </summary>
        partyGlass,

        /// <summary>
        /// Pencils Art Border
        /// </summary>
        pencils,

        /// <summary>
        /// Character Art Border
        /// </summary>
        people,

        /// <summary>
        /// Waving Character Border
        /// </summary>
        peopleWaving,

        /// <summary>
        /// Character With Hat Art Border
        /// </summary>
        peopleHats,

        /// <summary>
        /// Poinsettia Art Border
        /// </summary>
        poinsettias,

        /// <summary>
        /// Postage Stamp Art Border
        /// </summary>
        postageStamp,

        /// <summary>
        /// Pumpkin Art Border
        /// </summary>
        pumpkin1,

        /// <summary>
        /// Push Pin Art Border
        /// </summary>
        pushPinNote2,

        /// <summary>
        /// Push Pin Art Border
        /// </summary>
        pushPinNote1,

        /// <summary>
        /// Pyramid Art Border
        /// </summary>
        pyramids,

        /// <summary>
        /// Pyramid Art Border
        /// </summary>
        pyramidsAbove,

        /// <summary>
        /// Quadrants Art Border
        /// </summary>
        quadrants,

        /// <summary>
        /// Rings Art Border
        /// </summary>
        rings,

        /// <summary>
        /// Safari Art Border
        /// </summary>
        safari,

        /// <summary>
        /// Saw tooth Art Border
        /// </summary>
        sawtooth,

        /// <summary>
        /// Gray Saw tooth Art Border
        /// </summary>
        sawtoothGray,

        /// <summary>
        /// Scared Cat Art Border
        /// </summary>
        scaredCat,

        /// <summary>
        /// Umbrella Art Border
        /// </summary>
        seattle,

        /// <summary>
        /// Shadowed Squares Art Border
        /// </summary>
        shadowedSquares,

        /// <summary>
        /// Shark Tooth Art Border
        /// </summary>
        sharksTeeth,

        /// <summary>
        /// Bird Tracks Art Border
        /// </summary>
        shorebirdTracks,

        /// <summary>
        /// Rocket Art Border
        /// </summary>
        skyrocket,

        /// <summary>
        /// Snowflake Art Border
        /// </summary>
        snowflakeFancy,

        /// <summary>
        /// Snowflake Art Border
        /// </summary>
        snowflakes,

        /// <summary>
        /// Sombrero Art Border
        /// </summary>
        sombrero,

        /// <summary>
        /// Southwest-themed Art Border
        /// </summary>
        southwest,

        /// <summary>
        /// Stars Art Border
        /// </summary>
        stars,

        /// <summary>
        /// Stars On Top Art Border
        /// </summary>
        starsTop,

        /// <summary>
        /// 3-D Stars Art Border
        /// </summary>
        stars3d,

        /// <summary>
        /// Stars Art Border
        /// </summary>
        starsBlack,

        /// <summary>
        /// Stars With Shadows Art Border
        /// </summary>
        starsShadowed,

        /// <summary>
        /// Sun Art Border
        /// </summary>
        sun,

        /// <summary>
        /// Whirligig Art Border
        /// </summary>
        swirligig,

        /// <summary>
        /// Torn Paper Art Border
        /// </summary>
        tornPaper,

        /// <summary>
        /// Black Torn Paper Art Border
        /// </summary>
        tornPaperBlack,

        /// <summary>
        /// Tree Art Border
        /// </summary>
        trees,

        /// <summary>
        /// Triangle Art Border
        /// </summary>
        triangleParty,

        /// <summary>
        /// Triangles Art Border
        /// </summary>
        triangles,

        /// <summary>
        /// Tribal Art Border One
        /// </summary>
        tribal1,

        /// <summary>
        /// Tribal Art Border Two
        /// </summary>
        tribal2,

        /// <summary>
        /// Tribal Art Border Three
        /// </summary>
        tribal3,

        /// <summary>
        /// Tribal Art Border Four
        /// </summary>
        tribal4,

        /// <summary>
        /// Tribal Art Border Five
        /// </summary>
        tribal5,

        /// <summary>
        /// Tribal Art Border Six
        /// </summary>
        tribal6,

        /// <summary>
        /// Twisted Lines Art Border
        /// </summary>
        twistedLines1,

        /// <summary>
        /// Twisted Lines Art Border
        /// </summary>
        twistedLines2,

        /// <summary>
        /// Vine Art Border
        /// </summary>
        vine,

        /// <summary>
        /// Wavy Line Art Border
        /// </summary>
        waveline,

        /// <summary>
        /// Weaving Angles Art Border
        /// </summary>
        weavingAngles,

        /// <summary>
        /// Weaving Braid Art Border
        /// </summary>
        weavingBraid,

        /// <summary>
        /// Weaving Ribbon Art Border
        /// </summary>
        weavingRibbon,

        /// <summary>
        /// Weaving Strips Art Border
        /// </summary>
        weavingStrips,

        /// <summary>
        /// White Flowers Art Border
        /// </summary>
        whiteFlowers,

        /// <summary>
        /// Woodwork Art Border
        /// </summary>
        woodwork,

        /// <summary>
        /// Crisscross Art Border
        /// </summary>
        xIllusions,

        /// <summary>
        /// Triangle Art Border
        /// </summary>
        zanyTriangles,

        /// <summary>
        /// Zigzag Art Border
        /// </summary>
        zigZag,

        /// <summary>
        /// Zigzag stitch
        /// </summary>
        zigZagStitch,
    }
}