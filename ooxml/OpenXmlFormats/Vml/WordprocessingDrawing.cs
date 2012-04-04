using System.Xml.Serialization;
using System;
namespace NPOI.OpenXmlFormats.Vml
{
    
    
    /// <remarks/>
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
        
        /// <remarks/>
        [XmlAttribute]
        public ST_BorderType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute(DataType="positiveInteger")]
        public string width {
            get {
                return this.widthField;
            }
            set {
                this.widthField = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public ST_BorderShadow shadow {
            get {
                return this.shadowField;
            }
            set {
                this.shadowField = value;
            }
        }
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_BorderType {
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        single,
        
        /// <remarks/>
        thick,
        
        /// <remarks/>
        @double,
        
        /// <remarks/>
        hairline,
        
        /// <remarks/>
        dot,
        
        /// <remarks/>
        dash,
        
        /// <remarks/>
        dotDash,
        
        /// <remarks/>
        dashDotDot,
        
        /// <remarks/>
        triple,
        
        /// <remarks/>
        thinThickSmall,
        
        /// <remarks/>
        thickThinSmall,
        
        /// <remarks/>
        thickBetweenThinSmall,
        
        /// <remarks/>
        thinThick,
        
        /// <remarks/>
        thickThin,
        
        /// <remarks/>
        thickBetweenThin,
        
        /// <remarks/>
        thinThickLarge,
        
        /// <remarks/>
        thickThinLarge,
        
        /// <remarks/>
        thickBetweenThinLarge,
        
        /// <remarks/>
        wave,
        
        /// <remarks/>
        doubleWave,
        
        /// <remarks/>
        dashedSmall,
        
        /// <remarks/>
        dashDotStroked,
        
        /// <remarks/>
        threeDEmboss,
        
        /// <remarks/>
        threeDEngrave,
        
        /// <remarks/>
        HTMLOutset,
        
        /// <remarks/>
        HTMLInset,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_BorderShadow {
        
        /// <remarks/>
        t,
        
        /// <remarks/>
        @true,
        
        /// <remarks/>
        f,
        
        /// <remarks/>
        @false,
    }
    
    /// <remarks/>
    [System.Serializable]
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
        
        /// <remarks/>
        [XmlAttribute]
        public ST_WrapType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool typeSpecified {
            get {
                return this.typeFieldSpecified;
            }
            set {
                this.typeFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public ST_WrapSide side {
            get {
                return this.sideField;
            }
            set {
                this.sideField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool sideSpecified {
            get {
                return this.sideFieldSpecified;
            }
            set {
                this.sideFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public ST_HorizontalAnchor anchorx {
            get {
                return this.anchorxField;
            }
            set {
                this.anchorxField = value;
            }
        }
        
        /// <remarks/>
        [XmlIgnore]
        public bool anchorxSpecified {
            get {
                return this.anchorxFieldSpecified;
            }
            set {
                this.anchorxFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [XmlAttribute]
        public ST_VerticalAnchor anchory {
            get {
                return this.anchoryField;
            }
            set {
                this.anchoryField = value;
            }
        }
        
        /// <remarks/>
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
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_WrapType {
        
        /// <remarks/>
        topAndBottom,
        
        /// <remarks/>
        square,
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        tight,
        
        /// <remarks/>
        through,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_WrapSide {
        
        /// <remarks/>
        both,
        
        /// <remarks/>
        left,
        
        /// <remarks/>
        right,
        
        /// <remarks/>
        largest,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_HorizontalAnchor {
        
        /// <remarks/>
        margin,
        
        /// <remarks/>
        page,
        
        /// <remarks/>
        text,
        
        /// <remarks/>
        @char,
    }
    
    /// <remarks/>
    [System.Serializable]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=false)]
    public enum ST_VerticalAnchor {
        
        /// <remarks/>
        margin,
        
        /// <remarks/>
        page,
        
        /// <remarks/>
        text,
        
        /// <remarks/>
        line,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace="urn:schemas-microsoft-com:office:word")]
    [XmlRoot(Namespace="urn:schemas-microsoft-com:office:word", IsNullable=true)]
    public partial class CT_AnchorLock {
    }
}
