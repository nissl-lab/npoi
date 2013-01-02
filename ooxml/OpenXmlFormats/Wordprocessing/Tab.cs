using System;
using System.Collections.Generic;
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
