using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Tabs
    {

        private List<CT_TabStop> tabField;

        public CT_Tabs()
        {
            this.tabField = new List<CT_TabStop>();
        }

        [System.Xml.Serialization.XmlElement("tab", Order = 0)]
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

    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TabTlc
    {

    
        none,

    
        dot,

    
        hyphen,

    
        underscore,

    
        heavy,

    
        middleDot,
    }
    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TabStop
    {

        private ST_TabJc valField;

        private ST_TabTlc leaderField;

        private bool leaderFieldSpecified;

        private string posField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
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

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PTab
    {

        private ST_PTabAlignment alignmentField;

        private ST_PTabRelativeTo relativeToField;

        private ST_PTabLeader leaderField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PTabAlignment
    {

    
        left,

    
        center,

    
        right,
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PTabRelativeTo
    {

    
        margin,

    
        indent,
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PTabLeader
    {

    
        none,

    
        dot,

    
        hyphen,

    
        underscore,

    
        middleDot,
    }

}
