using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Tabs
    {

        private List<CT_TabStop> tabField;

        public CT_Tabs()
        {
            this.tabField = new List<CT_TabStop>();
        }

        [System.Xml.Serialization.XmlElementAttribute("tab", Order = 0)]
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
            throw new NotImplementedException();
        }
    }

    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TabJc
    {

        /// <remarks/>
        clear,

        /// <remarks/>
        left,

        /// <remarks/>
        center,

        /// <remarks/>
        right,

        /// <remarks/>
        @decimal,

        /// <remarks/>
        bar,

        /// <remarks/>
        num,
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_TabTlc
    {

        /// <remarks/>
        none,

        /// <remarks/>
        dot,

        /// <remarks/>
        hyphen,

        /// <remarks/>
        underscore,

        /// <remarks/>
        heavy,

        /// <remarks/>
        middleDot,
    }
    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
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

    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
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
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PTabAlignment
    {

        /// <remarks/>
        left,

        /// <remarks/>
        center,

        /// <remarks/>
        right,
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PTabRelativeTo
    {

        /// <remarks/>
        margin,

        /// <remarks/>
        indent,
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PTabLeader
    {

        /// <remarks/>
        none,

        /// <remarks/>
        dot,

        /// <remarks/>
        hyphen,

        /// <remarks/>
        underscore,

        /// <remarks/>
        middleDot,
    }

}
