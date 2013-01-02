using System;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Vml.Wordprocessing
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=true)]
    public partial class CT_Border {
        
        private ST_BorderType typeField;
        
        private bool typeFieldSpecified;
        
        private string widthField;
        
        private ST_BorderShadow shadowField;
        
        private bool shadowFieldSpecified;
        
    
        [XmlAttribute]
        public ST_BorderType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute(DataType="positiveInteger")]
        public string width {
            get {
                return this.widthField;
            }
            set {
                this.widthField = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_BorderShadow shadow {
            get {
                return this.shadowField;
            }
            set {
                this.shadowField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool shadowSpecified {
            get {
                return this.shadowFieldSpecified;
            }
            set {
                this.shadowFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_BorderType {
        
    
        none,
        
    
        single,
        
    
        thick,
        
    
        @double,
        
    
        hairline,
        
    
        dot,
        
    
        dash,
        
    
        dotDash,
        
    
        dashDotDot,
        
    
        triple,
        
    
        thinThickSmall,
        
    
        thickThinSmall,
        
    
        thickBetweenThinSmall,
        
    
        thinThick,
        
    
        thickThin,
        
    
        thickBetweenThin,
        
    
        thinThickLarge,
        
    
        thickThinLarge,
        
    
        thickBetweenThinLarge,
        
    
        wave,
        
    
        doubleWave,
        
    
        dashedSmall,
        
    
        dashDotStroked,
        
    
        threeDEmboss,
        
    
        threeDEngrave,
        
    
        HTMLOutset,
        
    
        HTMLInset,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_BorderShadow {
        
    
        t,
        
    
        @true,
        
    
        f,
        
    
        @false,
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=true)]
    public partial class CT_Wrap {
        
        private ST_WrapType typeField;
        
        private bool typeFieldSpecified;
        
        private ST_WrapSide sideField;
        
        private bool sideFieldSpecified;
        
        private ST_HorizontalAnchor anchorxField;
        
        private bool anchorxFieldSpecified;
        
        private ST_VerticalAnchor anchoryField;
        
        private bool anchoryFieldSpecified;
        
    
        [XmlAttribute]
        public ST_WrapType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_WrapSide side {
            get {
                return this.sideField;
            }
            set {
                this.sideField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool sideSpecified {
            get {
                return this.sideFieldSpecified;
            }
            set {
                this.sideFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_HorizontalAnchor anchorx {
            get {
                return this.anchorxField;
            }
            set {
                this.anchorxField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool anchorxSpecified {
            get {
                return this.anchorxFieldSpecified;
            }
            set {
                this.anchorxFieldSpecified = value;
            }
        }
        
    
        [XmlAttribute]
        public ST_VerticalAnchor anchory {
            get {
                return this.anchoryField;
            }
            set {
                this.anchoryField = value;
            }
        }
        
    
        [XmlIgnore]
        public bool anchorySpecified {
            get {
                return this.anchoryFieldSpecified;
            }
            set {
                this.anchoryFieldSpecified = value;
            }
        }
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_WrapType {
        
    
        topAndBottom,
        
    
        square,
        
    
        none,
        
    
        tight,
        
    
        through,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_WrapSide {
        
    
        both,
        
    
        left,
        
    
        right,
        
    
        largest,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_HorizontalAnchor {
        
    
        margin,
        
    
        page,
        
    
        text,
        
    
        @char,
    }
    

    [Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_VerticalAnchor {
        
    
        margin,
        
    
        page,
        
    
        text,
        
    
        line,
    }
    

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=true)]
    public partial class CT_AnchorLock {
    }
}
