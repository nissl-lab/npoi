using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXml4Net.Util;


namespace NPOI.OpenXmlFormats.Wordprocessing
{
    #region Range Permission

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_ProofErr
    {

    
        spellStart,

    
        spellEnd,

    
        gramStart,

    
        gramEnd,
    }

    [XmlInclude(typeof(CT_PermStart))]

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Perm
    {

        private string idField;

        private ST_DisplacedByCustomXml displacedByCustomXmlField;

        private bool displacedByCustomXmlFieldSpecified;
        public static CT_Perm Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Perm ctObj = new CT_Perm();
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            if (node.Attributes["w:displacedByCustomXml"] != null)
                ctObj.displacedByCustomXml = (ST_DisplacedByCustomXml)Enum.Parse(typeof(ST_DisplacedByCustomXml), node.Attributes["w:displacedByCustomXml"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            XmlHelper.WriteAttribute(sw, "w:displacedByCustomXml", this.displacedByCustomXml.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_DisplacedByCustomXml displacedByCustomXml
        {
            get
            {
                return this.displacedByCustomXmlField;
            }
            set
            {
                this.displacedByCustomXmlField = value;
            }
        }

        [XmlIgnore]
        public bool displacedByCustomXmlSpecified
        {
            get
            {
                return this.displacedByCustomXmlFieldSpecified;
            }
            set
            {
                this.displacedByCustomXmlFieldSpecified = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PermStart : CT_Perm
    {

        private ST_EdGrp edGrpField;

        private bool edGrpFieldSpecified;

        private string edField;

        private string colFirstField;

        private string colLastField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_EdGrp edGrp
        {
            get
            {
                return this.edGrpField;
            }
            set
            {
                this.edGrpField = value;
            }
        }

        [XmlIgnore]
        public bool edGrpSpecified
        {
            get
            {
                return this.edGrpFieldSpecified;
            }
            set
            {
                this.edGrpFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string ed
        {
            get
            {
                return this.edField;
            }
            set
            {
                this.edField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string colFirst
        {
            get
            {
                return this.colFirstField;
            }
            set
            {
                this.colFirstField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string colLast
        {
            get
            {
                return this.colLastField;
            }
            set
            {
                this.colLastField = value;
            }
        }
        public static new CT_PermStart Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PermStart ctObj = new CT_PermStart();
            if (node.Attributes["w:edGrp"] != null)
                ctObj.edGrp = (ST_EdGrp)Enum.Parse(typeof(ST_EdGrp), node.Attributes["w:edGrp"].Value);
            ctObj.ed = XmlHelper.ReadString(node.Attributes["w:ed"]);
            ctObj.colFirst = XmlHelper.ReadString(node.Attributes["w:colFirst"]);
            ctObj.colLast = XmlHelper.ReadString(node.Attributes["w:colLast"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["w:id"]);
            if (node.Attributes["w:displacedByCustomXml"] != null)
                ctObj.displacedByCustomXml = (ST_DisplacedByCustomXml)Enum.Parse(typeof(ST_DisplacedByCustomXml), node.Attributes["w:displacedByCustomXml"].Value);
            return ctObj;
        }



        internal new void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:edGrp", this.edGrp.ToString());
            XmlHelper.WriteAttribute(sw, "w:ed", this.ed);
            XmlHelper.WriteAttribute(sw, "w:colFirst", this.colFirst);
            XmlHelper.WriteAttribute(sw, "w:colLast", this.colLast);
            XmlHelper.WriteAttribute(sw, "w:id", this.id);
            XmlHelper.WriteAttribute(sw, "w:displacedByCustomXml", this.displacedByCustomXml.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_EdGrp
    {

    
        none,

    
        everyone,

    
        administrators,

    
        contributors,

    
        editors,

    
        owners,

    
        current,
    }
    #endregion

}
