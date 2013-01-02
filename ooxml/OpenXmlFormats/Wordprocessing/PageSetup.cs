using System;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PageSz
    {

        private ulong wField;

        private bool wFieldSpecified;

        private ulong hField;

        private bool hFieldSpecified;

        private ST_PageOrientation orientField;

        private bool orientFieldSpecified;

        private string codeField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong w
        {
            get
            {
                return this.wField;
            }
            set
            {
                this.wField = value;
            }
        }

        [XmlIgnore]
        public bool wSpecified
        {
            get
            {
                return this.wFieldSpecified;
            }
            set
            {
                this.wFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong h
        {
            get
            {
                return this.hField;
            }
            set
            {
                this.hField = value;
            }
        }

        [XmlIgnore]
        public bool hSpecified
        {
            get
            {
                return this.hFieldSpecified;
            }
            set
            {
                this.hFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PageOrientation orient
        {
            get
            {
                return this.orientField;
            }
            set
            {
                this.orientField = value;
            }
        }

        [XmlIgnore]
        public bool orientSpecified
        {
            get
            {
                return this.orientFieldSpecified;
            }
            set
            {
                this.orientFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string code
        {
            get
            {
                return this.codeField;
            }
            set
            {
                this.codeField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PageOrientation
    {

    
        portrait,

    
        landscape,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PageMar
    {

        private string topField;

        private ulong rightField;

        private string bottomField;

        private ulong leftField;

        private ulong headerField;

        private ulong footerField;

        private ulong gutterField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string top
        {
            get
            {
                return this.topField;
            }
            set
            {
                this.topField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string bottom
        {
            get
            {
                return this.bottomField;
            }
            set
            {
                this.bottomField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong header
        {
            get
            {
                return this.headerField;
            }
            set
            {
                this.headerField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong footer
        {
            get
            {
                return this.footerField;
            }
            set
            {
                this.footerField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong gutter
        {
            get
            {
                return this.gutterField;
            }
            set
            {
                this.gutterField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PaperSource
    {

        private string firstField;

        private string otherField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string first
        {
            get
            {
                return this.firstField;
            }
            set
            {
                this.firstField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string other
        {
            get
            {
                return this.otherField;
            }
            set
            {
                this.otherField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PageBorders
    {

        private CT_Border topField;

        private CT_Border leftField;

        private CT_Border bottomField;

        private CT_Border rightField;

        private ST_PageBorderZOrder zOrderField;

        private bool zOrderFieldSpecified;

        private ST_PageBorderDisplay displayField;

        private bool displayFieldSpecified;

        private ST_PageBorderOffset offsetFromField;

        private bool offsetFromFieldSpecified;

        public CT_PageBorders()
        {
            this.rightField = new CT_Border();
            this.bottomField = new CT_Border();
            this.leftField = new CT_Border();
            this.topField = new CT_Border();
        }

        [XmlElement(Order = 0)]
        public CT_Border top
        {
            get
            {
                return this.topField;
            }
            set
            {
                this.topField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Border left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_Border bottom
        {
            get
            {
                return this.bottomField;
            }
            set
            {
                this.bottomField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_Border right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PageBorderZOrder zOrder
        {
            get
            {
                return this.zOrderField;
            }
            set
            {
                this.zOrderField = value;
            }
        }

        [XmlIgnore]
        public bool zOrderSpecified
        {
            get
            {
                return this.zOrderFieldSpecified;
            }
            set
            {
                this.zOrderFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PageBorderDisplay display
        {
            get
            {
                return this.displayField;
            }
            set
            {
                this.displayField = value;
            }
        }

        [XmlIgnore]
        public bool displaySpecified
        {
            get
            {
                return this.displayFieldSpecified;
            }
            set
            {
                this.displayFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PageBorderOffset offsetFrom
        {
            get
            {
                return this.offsetFromField;
            }
            set
            {
                this.offsetFromField = value;
            }
        }

        [XmlIgnore]
        public bool offsetFromSpecified
        {
            get
            {
                return this.offsetFromFieldSpecified;
            }
            set
            {
                this.offsetFromFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PageBorderZOrder
    {

    
        front,

    
        back,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PageBorderDisplay
    {

    
        allPages,

    
        firstPage,

    
        notFirstPage,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PageBorderOffset
    {

    
        page,

    
        text,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PageNumber
    {

        private ST_NumberFormat fmtField;

        private bool fmtFieldSpecified;

        private string startField;

        private string chapStyleField;

        private ST_ChapterSep chapSepField;

        private bool chapSepFieldSpecified;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_NumberFormat fmt
        {
            get
            {
                return this.fmtField;
            }
            set
            {
                this.fmtField = value;
            }
        }

        [XmlIgnore]
        public bool fmtSpecified
        {
            get
            {
                return this.fmtFieldSpecified;
            }
            set
            {
                this.fmtFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string start
        {
            get
            {
                return this.startField;
            }
            set
            {
                this.startField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string chapStyle
        {
            get
            {
                return this.chapStyleField;
            }
            set
            {
                this.chapStyleField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ChapterSep chapSep
        {
            get
            {
                return this.chapSepField;
            }
            set
            {
                this.chapSepField = value;
            }
        }

        [XmlIgnore]
        public bool chapSepSpecified
        {
            get
            {
                return this.chapSepFieldSpecified;
            }
            set
            {
                this.chapSepFieldSpecified = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SectType
    {

        private ST_SectionMark valField;

        private bool valFieldSpecified;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_SectionMark val
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
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_SectionMark
    {

    
        nextPage,

    
        nextColumn,

    
        continuous,

    
        evenPage,

    
        oddPage,
    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LineNumber
    {

        private string countByField;

        private string startField;

        private ulong distanceField;

        private bool distanceFieldSpecified;

        private ST_LineNumberRestart restartField;

        private bool restartFieldSpecified;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string countBy
        {
            get
            {
                return this.countByField;
            }
            set
            {
                this.countByField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string start
        {
            get
            {
                return this.startField;
            }
            set
            {
                this.startField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong distance
        {
            get
            {
                return this.distanceField;
            }
            set
            {
                this.distanceField = value;
            }
        }

        [XmlIgnore]
        public bool distanceSpecified
        {
            get
            {
                return this.distanceFieldSpecified;
            }
            set
            {
                this.distanceFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_LineNumberRestart restart
        {
            get
            {
                return this.restartField;
            }
            set
            {
                this.restartField = value;
            }
        }

        [XmlIgnore]
        public bool restartSpecified
        {
            get
            {
                return this.restartFieldSpecified;
            }
            set
            {
                this.restartFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_LineNumberRestart
    {

    
        newPage,

    
        newSection,

    
        continuous,
    }
}
