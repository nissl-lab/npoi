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
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Tabs
    {

        private List<CT_TabStop> tabField;

        public CT_Tabs()
        {
            this.tabField = new List<CT_TabStop>();
        }

        [XmlElement("tab", Order = 0)]
        public List<CT_TabStop> tab
        {
            get
            {
                return this.tabField;
            }
            set
            {
                this.tabField = value;
            }
        }

        public CT_TabStop AddNewTab()
        {
            CT_TabStop s = new CT_TabStop();
            this.tabField.Add(s);
            return s;
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TabJc
    {

    
        clear,

    
        left,

    
        center,

    
        right,

    
        @decimal,

    
        bar,

    
        num,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TabTlc
    {

    
        none,

    
        dot,

    
        hyphen,

    
        underscore,

    
        heavy,

    
        middleDot,
    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TabStop
    {

        private ST_TabJc valField;

        private ST_TabTlc leaderField;

        private bool leaderFieldSpecified;

        private string posField;
        public static CT_TabStop Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TabStop ctObj = new CT_TabStop();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_TabJc)Enum.Parse(typeof(ST_TabJc), node.Attributes["w:val"].Value);
            if (node.Attributes["w:leader"] != null)
                ctObj.leader = (ST_TabTlc)Enum.Parse(typeof(ST_TabTlc), node.Attributes["w:leader"].Value);
            ctObj.pos = XmlHelper.ReadString(node.Attributes["w:pos"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            if(this.leader!= ST_TabTlc.none)
                XmlHelper.WriteAttribute(sw, "w:leader", this.leader.ToString());
            XmlHelper.WriteAttribute(sw, "w:pos", this.pos, true);
            sw.Write("/>");
        }


        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_TabJc val
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
        public ST_TabTlc leader
        {
            get
            {
                return this.leaderField;
            }
            set
            {
                this.leaderField = value;
            }
        }

        [XmlIgnore]
        public bool leaderSpecified
        {
            get
            {
                return this.leaderFieldSpecified;
            }
            set
            {
                this.leaderFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string pos
        {
            get
            {
                return this.posField;
            }
            set
            {
                this.posField = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PTab
    {

        private ST_PTabAlignment alignmentField;

        private ST_PTabRelativeTo relativeToField;

        private ST_PTabLeader leaderField;
        public static CT_PTab Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PTab ctObj = new CT_PTab();
            if (node.Attributes["w:alignment"] != null)
                ctObj.alignment = (ST_PTabAlignment)Enum.Parse(typeof(ST_PTabAlignment), node.Attributes["w:alignment"].Value);
            if (node.Attributes["w:relativeTo"] != null)
                ctObj.relativeTo = (ST_PTabRelativeTo)Enum.Parse(typeof(ST_PTabRelativeTo), node.Attributes["w:relativeTo"].Value);
            if (node.Attributes["w:leader"] != null)
                ctObj.leader = (ST_PTabLeader)Enum.Parse(typeof(ST_PTabLeader), node.Attributes["w:leader"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:alignment", this.alignment.ToString());
            XmlHelper.WriteAttribute(sw, "w:relativeTo", this.relativeTo.ToString());
            XmlHelper.WriteAttribute(sw, "w:leader", this.leader.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PTabAlignment alignment
        {
            get
            {
                return this.alignmentField;
            }
            set
            {
                this.alignmentField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PTabRelativeTo relativeTo
        {
            get
            {
                return this.relativeToField;
            }
            set
            {
                this.relativeToField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PTabLeader leader
        {
            get
            {
                return this.leaderField;
            }
            set
            {
                this.leaderField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PTabAlignment
    {

    
        left,

    
        center,

    
        right,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PTabRelativeTo
    {

    
        margin,

    
        indent,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PTabLeader
    {

    
        none,

    
        dot,

    
        hyphen,

    
        underscore,

    
        middleDot,
    }

}
