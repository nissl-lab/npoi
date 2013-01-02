using System;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Frameset
    {

        private CT_String szField;

        private CT_FramesetSplitbar framesetSplitbarField;

        private CT_FrameLayout frameLayoutField;

        private object[] itemsField;

        public CT_Frameset()
        {
            this.itemsField = new object[0];
            this.frameLayoutField = new CT_FrameLayout();
            this.framesetSplitbarField = new CT_FramesetSplitbar();
            this.szField = new CT_String();
        }

        [XmlElement(Order = 0)]
        public CT_String sz
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

        [XmlElement(Order = 1)]
        public CT_FramesetSplitbar framesetSplitbar
        {
            get
            {
                return this.framesetSplitbarField;
            }
            set
            {
                this.framesetSplitbarField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_FrameLayout frameLayout
        {
            get
            {
                return this.frameLayoutField;
            }
            set
            {
                this.frameLayoutField = value;
            }
        }

        [XmlElement("frame", typeof(CT_Frame), Order = 3)]
        [XmlElement("frameset", typeof(CT_Frameset), Order = 3)]
        public object[] Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FramesetSplitbar
    {

        private CT_TwipsMeasure wField;

        private CT_Color colorField;

        private CT_OnOff noBorderField;

        private CT_OnOff flatBordersField;

        public CT_FramesetSplitbar()
        {
            this.flatBordersField = new CT_OnOff();
            this.noBorderField = new CT_OnOff();
            this.colorField = new CT_Color();
            this.wField = new CT_TwipsMeasure();
        }

        [XmlElement(Order = 0)]
        public CT_TwipsMeasure w
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

        [XmlElement(Order = 1)]
        public CT_Color color
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

        [XmlElement(Order = 2)]
        public CT_OnOff noBorder
        {
            get
            {
                return this.noBorderField;
            }
            set
            {
                this.noBorderField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_OnOff flatBorders
        {
            get
            {
                return this.flatBordersField;
            }
            set
            {
                this.flatBordersField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FrameLayout
    {

        private ST_FrameLayout valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_FrameLayout val
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
    public enum ST_FrameLayout
    {

    
        rows,

    
        cols,

    
        none,
    }

    
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Frame
    {

        private CT_String szField;

        private CT_String nameField;

        private CT_Rel sourceFileNameField;

        private CT_PixelsMeasure marWField;

        private CT_PixelsMeasure marHField;

        private CT_FrameScrollbar scrollbarField;

        private CT_OnOff noResizeAllowedField;

        private CT_OnOff linkedToFileField;

        public CT_Frame()
        {
            this.linkedToFileField = new CT_OnOff();
            this.noResizeAllowedField = new CT_OnOff();
            this.scrollbarField = new CT_FrameScrollbar();
            this.marHField = new CT_PixelsMeasure();
            this.marWField = new CT_PixelsMeasure();
            this.sourceFileNameField = new CT_Rel();
            this.nameField = new CT_String();
            this.szField = new CT_String();
        }

        [XmlElement(Order = 0)]
        public CT_String sz
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
        public CT_Rel sourceFileName
        {
            get
            {
                return this.sourceFileNameField;
            }
            set
            {
                this.sourceFileNameField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_PixelsMeasure marW
        {
            get
            {
                return this.marWField;
            }
            set
            {
                this.marWField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_PixelsMeasure marH
        {
            get
            {
                return this.marHField;
            }
            set
            {
                this.marHField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_FrameScrollbar scrollbar
        {
            get
            {
                return this.scrollbarField;
            }
            set
            {
                this.scrollbarField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_OnOff noResizeAllowed
        {
            get
            {
                return this.noResizeAllowedField;
            }
            set
            {
                this.noResizeAllowedField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_OnOff linkedToFile
        {
            get
            {
                return this.linkedToFileField;
            }
            set
            {
                this.linkedToFileField = value;
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FrameScrollbar
    {

        private ST_FrameScrollbar valField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_FrameScrollbar val
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
    public enum ST_FrameScrollbar
    {

    
        on,

    
        off,

    
        auto,
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FramePr
    {

        private ST_DropCap dropCapField;

        private bool dropCapFieldSpecified;

        private string linesField;

        private ulong wField;

        private bool wFieldSpecified;

        private ulong hField;

        private bool hFieldSpecified;

        private ulong vSpaceField;

        private bool vSpaceFieldSpecified;

        private ulong hSpaceField;

        private bool hSpaceFieldSpecified;

        private ST_Wrap wrapField;

        private bool wrapFieldSpecified;

        private ST_HAnchor hAnchorField;

        private bool hAnchorFieldSpecified;

        private ST_VAnchor vAnchorField;

        private bool vAnchorFieldSpecified;

        private string xField;

        private ST_XAlign xAlignField;

        private bool xAlignFieldSpecified;

        private string yField;

        private ST_YAlign yAlignField;

        private bool yAlignFieldSpecified;

        private ST_HeightRule hRuleField;

        private bool hRuleFieldSpecified;

        private ST_OnOff anchorLockField;

        private bool anchorLockFieldSpecified;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_DropCap dropCap
        {
            get
            {
                return this.dropCapField;
            }
            set
            {
                this.dropCapField = value;
            }
        }

        [XmlIgnore]
        public bool dropCapSpecified
        {
            get
            {
                return this.dropCapFieldSpecified;
            }
            set
            {
                this.dropCapFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string lines
        {
            get
            {
                return this.linesField;
            }
            set
            {
                this.linesField = value;
            }
        }

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
        public ulong vSpace
        {
            get
            {
                return this.vSpaceField;
            }
            set
            {
                this.vSpaceField = value;
            }
        }

        [XmlIgnore]
        public bool vSpaceSpecified
        {
            get
            {
                return this.vSpaceFieldSpecified;
            }
            set
            {
                this.vSpaceFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong hSpace
        {
            get
            {
                return this.hSpaceField;
            }
            set
            {
                this.hSpaceField = value;
            }
        }

        [XmlIgnore]
        public bool hSpaceSpecified
        {
            get
            {
                return this.hSpaceFieldSpecified;
            }
            set
            {
                this.hSpaceFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_Wrap wrap
        {
            get
            {
                return this.wrapField;
            }
            set
            {
                this.wrapField = value;
            }
        }

        [XmlIgnore]
        public bool wrapSpecified
        {
            get
            {
                return this.wrapFieldSpecified;
            }
            set
            {
                this.wrapFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_HAnchor hAnchor
        {
            get
            {
                return this.hAnchorField;
            }
            set
            {
                this.hAnchorField = value;
            }
        }

        [XmlIgnore]
        public bool hAnchorSpecified
        {
            get
            {
                return this.hAnchorFieldSpecified;
            }
            set
            {
                this.hAnchorFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_VAnchor vAnchor
        {
            get
            {
                return this.vAnchorField;
            }
            set
            {
                this.vAnchorField = value;
            }
        }

        [XmlIgnore]
        public bool vAnchorSpecified
        {
            get
            {
                return this.vAnchorFieldSpecified;
            }
            set
            {
                this.vAnchorFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string x
        {
            get
            {
                return this.xField;
            }
            set
            {
                this.xField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_XAlign xAlign
        {
            get
            {
                return this.xAlignField;
            }
            set
            {
                this.xAlignField = value;
            }
        }

        [XmlIgnore]
        public bool xAlignSpecified
        {
            get
            {
                return this.xAlignFieldSpecified;
            }
            set
            {
                this.xAlignFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string y
        {
            get
            {
                return this.yField;
            }
            set
            {
                this.yField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_YAlign yAlign
        {
            get
            {
                return this.yAlignField;
            }
            set
            {
                this.yAlignField = value;
            }
        }

        [XmlIgnore]
        public bool yAlignSpecified
        {
            get
            {
                return this.yAlignFieldSpecified;
            }
            set
            {
                this.yAlignFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_HeightRule hRule
        {
            get
            {
                return this.hRuleField;
            }
            set
            {
                this.hRuleField = value;
            }
        }

        [XmlIgnore]
        public bool hRuleSpecified
        {
            get
            {
                return this.hRuleFieldSpecified;
            }
            set
            {
                this.hRuleFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff anchorLock
        {
            get
            {
                return this.anchorLockField;
            }
            set
            {
                this.anchorLockField = value;
            }
        }

        [XmlIgnore]
        public bool anchorLockSpecified
        {
            get
            {
                return this.anchorLockFieldSpecified;
            }
            set
            {
                this.anchorLockFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_DropCap
    {

    
        none,

    
        drop,

    
        margin,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_Wrap
    {

    
        auto,

    
        notBeside,

    
        around,

    
        tight,

    
        through,

    
        none,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_HAnchor
    {

    
        text,

    
        margin,

    
        page,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_VAnchor
    {

    
        text,

    
        margin,

    
        page,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_XAlign
    {

    
        left,

    
        center,

    
        right,

    
        inside,

    
        outside,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_YAlign
    {

    
        inline,

    
        top,

    
        center,

    
        bottom,

    
        inside,

    
        outside,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_HeightRule
    {

    
        auto,

    
        exact,

    
        atLeast,
    }
}