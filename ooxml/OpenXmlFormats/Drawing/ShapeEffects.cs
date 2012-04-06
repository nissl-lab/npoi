using System.ComponentModel;
namespace NPOI.OpenXmlFormats.Dml
{
    
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AlphaBiLevelEffect {
        
        private int threshField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int thresh {
            get {
                return this.threshField;
            }
            set {
                this.threshField = value;
            }
        }
    }
    
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_HslColor {
        
        private CT_PositiveFixedPercentage[] tintField;
        
        private CT_PositiveFixedPercentage[] shadeField;
        
        private CT_ComplementTransform[] compField;
        
        private CT_InverseTransform[] invField;
        
        private CT_GrayscaleTransform[] grayField;
        
        private CT_PositiveFixedPercentage[] alphaField;
        
        private CT_FixedPercentage[] alphaOffField;
        
        private CT_PositivePercentage[] alphaModField;
        
        private CT_PositiveFixedAngle[] hueField;
        
        private CT_Angle[] hueOffField;
        
        private CT_PositivePercentage[] hueModField;
        
        private CT_Percentage[] satField;
        
        private CT_Percentage[] satOffField;
        
        private CT_Percentage[] satModField;
        
        private CT_Percentage[] lumField;
        
        private CT_Percentage[] lumOffField;
        
        private CT_Percentage[] lumModField;
        
        private CT_Percentage[] redField;
        
        private CT_Percentage[] redOffField;
        
        private CT_Percentage[] redModField;
        
        private CT_Percentage[] greenField;
        
        private CT_Percentage[] greenOffField;
        
        private CT_Percentage[] greenModField;
        
        private CT_Percentage[] blueField;
        
        private CT_Percentage[] blueOffField;
        
        private CT_Percentage[] blueModField;
        
        private CT_GammaTransform[] gammaField;
        
        private CT_InverseGammaTransform[] invGammaField;
        
        private int hue1Field;
        
        private int sat1Field;
        
        private int lum1Field;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("tint")]
        public CT_PositiveFixedPercentage[] tint {
            get {
                return this.tintField;
            }
            set {
                this.tintField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("shade")]
        public CT_PositiveFixedPercentage[] shade {
            get {
                return this.shadeField;
            }
            set {
                this.shadeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("comp")]
        public CT_ComplementTransform[] comp {
            get {
                return this.compField;
            }
            set {
                this.compField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("inv")]
        public CT_InverseTransform[] inv {
            get {
                return this.invField;
            }
            set {
                this.invField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("gray")]
        public CT_GrayscaleTransform[] gray {
            get {
                return this.grayField;
            }
            set {
                this.grayField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("alpha")]
        public CT_PositiveFixedPercentage[] alpha {
            get {
                return this.alphaField;
            }
            set {
                this.alphaField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("alphaOff")]
        public CT_FixedPercentage[] alphaOff {
            get {
                return this.alphaOffField;
            }
            set {
                this.alphaOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("alphaMod")]
        public CT_PositivePercentage[] alphaMod {
            get {
                return this.alphaModField;
            }
            set {
                this.alphaModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("hue")]
        public CT_PositiveFixedAngle[] hue {
            get {
                return this.hueField;
            }
            set {
                this.hueField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("hueOff")]
        public CT_Angle[] hueOff {
            get {
                return this.hueOffField;
            }
            set {
                this.hueOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("hueMod")]
        public CT_PositivePercentage[] hueMod {
            get {
                return this.hueModField;
            }
            set {
                this.hueModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("sat")]
        public CT_Percentage[] sat {
            get {
                return this.satField;
            }
            set {
                this.satField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("satOff")]
        public CT_Percentage[] satOff {
            get {
                return this.satOffField;
            }
            set {
                this.satOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("satMod")]
        public CT_Percentage[] satMod {
            get {
                return this.satModField;
            }
            set {
                this.satModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("lum")]
        public CT_Percentage[] lum {
            get {
                return this.lumField;
            }
            set {
                this.lumField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("lumOff")]
        public CT_Percentage[] lumOff {
            get {
                return this.lumOffField;
            }
            set {
                this.lumOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("lumMod")]
        public CT_Percentage[] lumMod {
            get {
                return this.lumModField;
            }
            set {
                this.lumModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("red")]
        public CT_Percentage[] red {
            get {
                return this.redField;
            }
            set {
                this.redField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("redOff")]
        public CT_Percentage[] redOff {
            get {
                return this.redOffField;
            }
            set {
                this.redOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("redMod")]
        public CT_Percentage[] redMod {
            get {
                return this.redModField;
            }
            set {
                this.redModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("green")]
        public CT_Percentage[] green {
            get {
                return this.greenField;
            }
            set {
                this.greenField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("greenOff")]
        public CT_Percentage[] greenOff {
            get {
                return this.greenOffField;
            }
            set {
                this.greenOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("greenMod")]
        public CT_Percentage[] greenMod {
            get {
                return this.greenModField;
            }
            set {
                this.greenModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("blue")]
        public CT_Percentage[] blue {
            get {
                return this.blueField;
            }
            set {
                this.blueField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("blueOff")]
        public CT_Percentage[] blueOff {
            get {
                return this.blueOffField;
            }
            set {
                this.blueOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("blueMod")]
        public CT_Percentage[] blueMod {
            get {
                return this.blueModField;
            }
            set {
                this.blueModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("gamma")]
        public CT_GammaTransform[] gamma {
            get {
                return this.gammaField;
            }
            set {
                this.gammaField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("invGamma")]
        public CT_InverseGammaTransform[] invGamma {
            get {
                return this.invGammaField;
            }
            set {
                this.invGammaField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute("hue")]
        public int hue1 {
            get {
                return this.hue1Field;
            }
            set {
                this.hue1Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute("sat")]
        public int sat1 {
            get {
                return this.sat1Field;
            }
            set {
                this.sat1Field = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute("lum")]
        public int lum1 {
            get {
                return this.lum1Field;
            }
            set {
                this.lum1Field = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_SystemColor {
        
        private CT_PositiveFixedPercentage[] tintField;
        
        private CT_PositiveFixedPercentage[] shadeField;
        
        private CT_ComplementTransform[] compField;
        
        private CT_InverseTransform[] invField;
        
        private CT_GrayscaleTransform[] grayField;
        
        private CT_PositiveFixedPercentage[] alphaField;
        
        private CT_FixedPercentage[] alphaOffField;
        
        private CT_PositivePercentage[] alphaModField;
        
        private CT_PositiveFixedAngle[] hueField;
        
        private CT_Angle[] hueOffField;
        
        private CT_PositivePercentage[] hueModField;
        
        private CT_Percentage[] satField;
        
        private CT_Percentage[] satOffField;
        
        private CT_Percentage[] satModField;
        
        private CT_Percentage[] lumField;
        
        private CT_Percentage[] lumOffField;
        
        private CT_Percentage[] lumModField;
        
        private CT_Percentage[] redField;
        
        private CT_Percentage[] redOffField;
        
        private CT_Percentage[] redModField;
        
        private CT_Percentage[] greenField;
        
        private CT_Percentage[] greenOffField;
        
        private CT_Percentage[] greenModField;
        
        private CT_Percentage[] blueField;
        
        private CT_Percentage[] blueOffField;
        
        private CT_Percentage[] blueModField;
        
        private CT_GammaTransform[] gammaField;
        
        private CT_InverseGammaTransform[] invGammaField;
        
        private ST_SystemColorVal valField;
        
        private byte[] lastClrField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("tint")]
        public CT_PositiveFixedPercentage[] tint {
            get {
                return this.tintField;
            }
            set {
                this.tintField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("shade")]
        public CT_PositiveFixedPercentage[] shade {
            get {
                return this.shadeField;
            }
            set {
                this.shadeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("comp")]
        public CT_ComplementTransform[] comp {
            get {
                return this.compField;
            }
            set {
                this.compField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("inv")]
        public CT_InverseTransform[] inv {
            get {
                return this.invField;
            }
            set {
                this.invField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("gray")]
        public CT_GrayscaleTransform[] gray {
            get {
                return this.grayField;
            }
            set {
                this.grayField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("alpha")]
        public CT_PositiveFixedPercentage[] alpha {
            get {
                return this.alphaField;
            }
            set {
                this.alphaField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("alphaOff")]
        public CT_FixedPercentage[] alphaOff {
            get {
                return this.alphaOffField;
            }
            set {
                this.alphaOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("alphaMod")]
        public CT_PositivePercentage[] alphaMod {
            get {
                return this.alphaModField;
            }
            set {
                this.alphaModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("hue")]
        public CT_PositiveFixedAngle[] hue {
            get {
                return this.hueField;
            }
            set {
                this.hueField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("hueOff")]
        public CT_Angle[] hueOff {
            get {
                return this.hueOffField;
            }
            set {
                this.hueOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("hueMod")]
        public CT_PositivePercentage[] hueMod {
            get {
                return this.hueModField;
            }
            set {
                this.hueModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("sat")]
        public CT_Percentage[] sat {
            get {
                return this.satField;
            }
            set {
                this.satField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("satOff")]
        public CT_Percentage[] satOff {
            get {
                return this.satOffField;
            }
            set {
                this.satOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("satMod")]
        public CT_Percentage[] satMod {
            get {
                return this.satModField;
            }
            set {
                this.satModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("lum")]
        public CT_Percentage[] lum {
            get {
                return this.lumField;
            }
            set {
                this.lumField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("lumOff")]
        public CT_Percentage[] lumOff {
            get {
                return this.lumOffField;
            }
            set {
                this.lumOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("lumMod")]
        public CT_Percentage[] lumMod {
            get {
                return this.lumModField;
            }
            set {
                this.lumModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("red")]
        public CT_Percentage[] red {
            get {
                return this.redField;
            }
            set {
                this.redField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("redOff")]
        public CT_Percentage[] redOff {
            get {
                return this.redOffField;
            }
            set {
                this.redOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("redMod")]
        public CT_Percentage[] redMod {
            get {
                return this.redModField;
            }
            set {
                this.redModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("green")]
        public CT_Percentage[] green {
            get {
                return this.greenField;
            }
            set {
                this.greenField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("greenOff")]
        public CT_Percentage[] greenOff {
            get {
                return this.greenOffField;
            }
            set {
                this.greenOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("greenMod")]
        public CT_Percentage[] greenMod {
            get {
                return this.greenModField;
            }
            set {
                this.greenModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("blue")]
        public CT_Percentage[] blue {
            get {
                return this.blueField;
            }
            set {
                this.blueField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("blueOff")]
        public CT_Percentage[] blueOff {
            get {
                return this.blueOffField;
            }
            set {
                this.blueOffField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("blueMod")]
        public CT_Percentage[] blueMod {
            get {
                return this.blueModField;
            }
            set {
                this.blueModField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("gamma")]
        public CT_GammaTransform[] gamma {
            get {
                return this.gammaField;
            }
            set {
                this.gammaField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("invGamma")]
        public CT_InverseGammaTransform[] invGamma {
            get {
                return this.invGammaField;
            }
            set {
                this.invGammaField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_SystemColorVal val {
            get {
                return this.valField;
            }
            set {
                this.valField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="hexBinary")]
        public byte[] lastClr {
            get {
                return this.lastClrField;
            }
            set {
                this.lastClrField = value;
            }
        }
    }
    
    
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AlphaCeilingEffect {
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AlphaFloorEffect {
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AlphaInverseEffect {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AlphaModulateFixedEffect {
        
        private int amtField;
        
        public CT_AlphaModulateFixedEffect() {
            this.amtField = 100000;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(100000)]
        public int amt {
            get {
                return this.amtField;
            }
            set {
                this.amtField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AlphaOutsetEffect {
        
        private long radField;
        
        public CT_AlphaOutsetEffect() {
            this.radField = ((long)(0));
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long rad {
            get {
                return this.radField;
            }
            set {
                this.radField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AlphaReplaceEffect {
        
        private int aField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int a {
            get {
                return this.aField;
            }
            set {
                this.aField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_BiLevelEffect {
        
        private int threshField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int thresh {
            get {
                return this.threshField;
            }
            set {
                this.threshField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_BlurEffect {
        
        private long radField;
        
        private bool growField;
        
        public CT_BlurEffect() {
            this.radField = ((long)(0));
            this.growField = true;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long rad {
            get {
                return this.radField;
            }
            set {
                this.radField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(true)]
        public bool grow {
            get {
                return this.growField;
            }
            set {
                this.growField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_ColorChangeEffect {
        
        private CT_Color clrFromField;
        
        private CT_Color clrToField;
        
        private bool useAField;
        
        public CT_ColorChangeEffect() {
            this.useAField = true;
        }
        
        /// <remarks/>
        public CT_Color clrFrom {
            get {
                return this.clrFromField;
            }
            set {
                this.clrFromField = value;
            }
        }
        
        /// <remarks/>
        public CT_Color clrTo {
            get {
                return this.clrToField;
            }
            set {
                this.clrToField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(true)]
        public bool useA {
            get {
                return this.useAField;
            }
            set {
                this.useAField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_ColorReplaceEffect {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_DuotoneEffect {
        
        private CT_ScRgbColor[] scrgbClrField;
        
        private CT_SRgbColor[] srgbClrField;
        
        private CT_HslColor[] hslClrField;
        
        private CT_SystemColor[] sysClrField;
        
        private CT_SchemeColor[] schemeClrField;
        
        private CT_PresetColor[] prstClrField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("scrgbClr")]
        public CT_ScRgbColor[] scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("srgbClr")]
        public CT_SRgbColor[] srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("hslClr")]
        public CT_HslColor[] hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("sysClr")]
        public CT_SystemColor[] sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("schemeClr")]
        public CT_SchemeColor[] schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("prstClr")]
        public CT_PresetColor[] prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GlowEffect {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        private long radField;
        
        public CT_GlowEffect() {
            this.radField = ((long)(0));
        }
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long rad {
            get {
                return this.radField;
            }
            set {
                this.radField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GrayscaleEffect {
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_HSLEffect {
        
        private int hueField;
        
        private int satField;
        
        private int lumField;
        
        public CT_HSLEffect() {
            this.hueField = 0;
            this.satField = 0;
            this.lumField = 0;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int hue {
            get {
                return this.hueField;
            }
            set {
                this.hueField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int sat {
            get {
                return this.satField;
            }
            set {
                this.satField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int lum {
            get {
                return this.lumField;
            }
            set {
                this.lumField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_InnerShadowEffect {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        private long blurRadField;
        
        private long distField;
        
        private int dirField;
        
        public CT_InnerShadowEffect() {
            this.blurRadField = ((long)(0));
            this.distField = ((long)(0));
            this.dirField = 0;
        }
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long blurRad {
            get {
                return this.blurRadField;
            }
            set {
                this.blurRadField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long dist {
            get {
                return this.distField;
            }
            set {
                this.distField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int dir {
            get {
                return this.dirField;
            }
            set {
                this.dirField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_LuminanceEffect {
        
        private int brightField;
        
        private int contrastField;
        
        public CT_LuminanceEffect() {
            this.brightField = 0;
            this.contrastField = 0;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int bright {
            get {
                return this.brightField;
            }
            set {
                this.brightField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int contrast {
            get {
                return this.contrastField;
            }
            set {
                this.contrastField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_OuterShadowEffect {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        private long blurRadField;
        
        private long distField;
        
        private int dirField;
        
        private int sxField;
        
        private int syField;
        
        private int kxField;
        
        private int kyField;
        
        private ST_RectAlignment algnField;
        
        private bool rotWithShapeField;
        
        public CT_OuterShadowEffect() {
            this.blurRadField = ((long)(0));
            this.distField = ((long)(0));
            this.dirField = 0;
            this.sxField = 100000;
            this.syField = 100000;
            this.kxField = 0;
            this.kyField = 0;
            this.algnField = ST_RectAlignment.b;
            this.rotWithShapeField = true;
        }
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long blurRad {
            get {
                return this.blurRadField;
            }
            set {
                this.blurRadField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long dist {
            get {
                return this.distField;
            }
            set {
                this.distField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int dir {
            get {
                return this.dirField;
            }
            set {
                this.dirField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(100000)]
        public int sx {
            get {
                return this.sxField;
            }
            set {
                this.sxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(100000)]
        public int sy {
            get {
                return this.syField;
            }
            set {
                this.syField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int kx {
            get {
                return this.kxField;
            }
            set {
                this.kxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int ky {
            get {
                return this.kyField;
            }
            set {
                this.kyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_RectAlignment.b)]
        public ST_RectAlignment algn {
            get {
                return this.algnField;
            }
            set {
                this.algnField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(true)]
        public bool rotWithShape {
            get {
                return this.rotWithShapeField;
            }
            set {
                this.rotWithShapeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_PresetShadowVal {
        
        /// <remarks/>
        shdw1,
        
        /// <remarks/>
        shdw2,
        
        /// <remarks/>
        shdw3,
        
        /// <remarks/>
        shdw4,
        
        /// <remarks/>
        shdw5,
        
        /// <remarks/>
        shdw6,
        
        /// <remarks/>
        shdw7,
        
        /// <remarks/>
        shdw8,
        
        /// <remarks/>
        shdw9,
        
        /// <remarks/>
        shdw10,
        
        /// <remarks/>
        shdw11,
        
        /// <remarks/>
        shdw12,
        
        /// <remarks/>
        shdw13,
        
        /// <remarks/>
        shdw14,
        
        /// <remarks/>
        shdw15,
        
        /// <remarks/>
        shdw16,
        
        /// <remarks/>
        shdw17,
        
        /// <remarks/>
        shdw18,
        
        /// <remarks/>
        shdw19,
        
        /// <remarks/>
        shdw20,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_PresetShadowEffect {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        private ST_PresetShadowVal prstField;
        
        private long distField;
        
        private int dirField;
        
        public CT_PresetShadowEffect() {
            this.distField = ((long)(0));
            this.dirField = 0;
        }
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_PresetShadowVal prst {
            get {
                return this.prstField;
            }
            set {
                this.prstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long dist {
            get {
                return this.distField;
            }
            set {
                this.distField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int dir {
            get {
                return this.dirField;
            }
            set {
                this.dirField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_ReflectionEffect {
        
        private long blurRadField;
        
        private int stAField;
        
        private int stPosField;
        
        private int endAField;
        
        private int endPosField;
        
        private long distField;
        
        private int dirField;
        
        private int fadeDirField;
        
        private int sxField;
        
        private int syField;
        
        private int kxField;
        
        private int kyField;
        
        private ST_RectAlignment algnField;
        
        private bool rotWithShapeField;
        
        public CT_ReflectionEffect() {
            this.blurRadField = ((long)(0));
            this.stAField = 100000;
            this.stPosField = 0;
            this.endAField = 0;
            this.endPosField = 100000;
            this.distField = ((long)(0));
            this.dirField = 0;
            this.fadeDirField = 5400000;
            this.sxField = 100000;
            this.syField = 100000;
            this.kxField = 0;
            this.kyField = 0;
            this.algnField = ST_RectAlignment.b;
            this.rotWithShapeField = true;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long blurRad {
            get {
                return this.blurRadField;
            }
            set {
                this.blurRadField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(100000)]
        public int stA {
            get {
                return this.stAField;
            }
            set {
                this.stAField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int stPos {
            get {
                return this.stPosField;
            }
            set {
                this.stPosField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int endA {
            get {
                return this.endAField;
            }
            set {
                this.endAField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(100000)]
        public int endPos {
            get {
                return this.endPosField;
            }
            set {
                this.endPosField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long dist {
            get {
                return this.distField;
            }
            set {
                this.distField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int dir {
            get {
                return this.dirField;
            }
            set {
                this.dirField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(5400000)]
        public int fadeDir {
            get {
                return this.fadeDirField;
            }
            set {
                this.fadeDirField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(100000)]
        public int sx {
            get {
                return this.sxField;
            }
            set {
                this.sxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(100000)]
        public int sy {
            get {
                return this.syField;
            }
            set {
                this.syField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int kx {
            get {
                return this.kxField;
            }
            set {
                this.kxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int ky {
            get {
                return this.kyField;
            }
            set {
                this.kyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_RectAlignment.b)]
        public ST_RectAlignment algn {
            get {
                return this.algnField;
            }
            set {
                this.algnField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(true)]
        public bool rotWithShape {
            get {
                return this.rotWithShapeField;
            }
            set {
                this.rotWithShapeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_RelativeOffsetEffect {
        
        private int txField;
        
        private int tyField;
        
        public CT_RelativeOffsetEffect() {
            this.txField = 0;
            this.tyField = 0;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int tx {
            get {
                return this.txField;
            }
            set {
                this.txField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int ty {
            get {
                return this.tyField;
            }
            set {
                this.tyField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_SoftEdgesEffect {
        
        private long radField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public long rad {
            get {
                return this.radField;
            }
            set {
                this.radField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TintEffect {
        
        private int hueField;
        
        private int amtField;
        
        public CT_TintEffect() {
            this.hueField = 0;
            this.amtField = 0;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int hue {
            get {
                return this.hueField;
            }
            set {
                this.hueField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int amt {
            get {
                return this.amtField;
            }
            set {
                this.amtField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TransformEffect {
        
        private int sxField;
        
        private int syField;
        
        private int kxField;
        
        private int kyField;
        
        private long txField;
        
        private long tyField;
        
        public CT_TransformEffect() {
            this.sxField = 100000;
            this.syField = 100000;
            this.kxField = 0;
            this.kyField = 0;
            this.txField = ((long)(0));
            this.tyField = ((long)(0));
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(100000)]
        public int sx {
            get {
                return this.sxField;
            }
            set {
                this.sxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(100000)]
        public int sy {
            get {
                return this.syField;
            }
            set {
                this.syField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int kx {
            get {
                return this.kxField;
            }
            set {
                this.kxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(0)]
        public int ky {
            get {
                return this.kyField;
            }
            set {
                this.kyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long tx {
            get {
                return this.txField;
            }
            set {
                this.txField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long ty {
            get {
                return this.tyField;
            }
            set {
                this.tyField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_NoFillProperties {
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_SolidColorFillProperties {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_LinearShadeProperties {
        
        private int angField;
        
        private bool angFieldSpecified;
        
        private bool scaledField;
        
        private bool scaledFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int ang {
            get {
                return this.angField;
            }
            set {
                this.angField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool angSpecified {
            get {
                return this.angFieldSpecified;
            }
            set {
                this.angFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public bool scaled {
            get {
                return this.scaledField;
            }
            set {
                this.scaledField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool scaledSpecified {
            get {
                return this.scaledFieldSpecified;
            }
            set {
                this.scaledFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_PathShadeType {
        
        /// <remarks/>
        shape,
        
        /// <remarks/>
        circle,
        
        /// <remarks/>
        rect,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_PathShadeProperties {
        
        private CT_RelativeRect fillToRectField;
        
        private ST_PathShadeType pathField;
        
        private bool pathFieldSpecified;
        
        /// <remarks/>
        public CT_RelativeRect fillToRect {
            get {
                return this.fillToRectField;
            }
            set {
                this.fillToRectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_PathShadeType path {
            get {
                return this.pathField;
            }
            set {
                this.pathField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool pathSpecified {
            get {
                return this.pathFieldSpecified;
            }
            set {
                this.pathFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_TileFlipMode {
        
        /// <remarks/>
        none,
        
        /// <remarks/>
        x,
        
        /// <remarks/>
        y,
        
        /// <remarks/>
        xy,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GradientStop {
        
        private CT_ScRgbColor scrgbClrField;
        
        private CT_SRgbColor srgbClrField;
        
        private CT_HslColor hslClrField;
        
        private CT_SystemColor sysClrField;
        
        private CT_SchemeColor schemeClrField;
        
        private CT_PresetColor prstClrField;
        
        private int posField;
        
        /// <remarks/>
        public CT_ScRgbColor scrgbClr {
            get {
                return this.scrgbClrField;
            }
            set {
                this.scrgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SRgbColor srgbClr {
            get {
                return this.srgbClrField;
            }
            set {
                this.srgbClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_HslColor hslClr {
            get {
                return this.hslClrField;
            }
            set {
                this.hslClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SystemColor sysClr {
            get {
                return this.sysClrField;
            }
            set {
                this.sysClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_SchemeColor schemeClr {
            get {
                return this.schemeClrField;
            }
            set {
                this.schemeClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetColor prstClr {
            get {
                return this.prstClrField;
            }
            set {
                this.prstClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int pos {
            get {
                return this.posField;
            }
            set {
                this.posField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GradientStopList {
        
        private CT_GradientStop[] gsField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("gs")]
        public CT_GradientStop[] gs {
            get {
                return this.gsField;
            }
            set {
                this.gsField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GradientFillProperties {
        
        private CT_GradientStop[] gsLstField;
        
        private CT_LinearShadeProperties linField;
        
        private CT_PathShadeProperties pathField;
        
        private CT_RelativeRect tileRectField;
        
        private ST_TileFlipMode flipField;
        
        private bool flipFieldSpecified;
        
        private bool rotWithShapeField;
        
        private bool rotWithShapeFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("gs", IsNullable=false)]
        public CT_GradientStop[] gsLst {
            get {
                return this.gsLstField;
            }
            set {
                this.gsLstField = value;
            }
        }
        
        /// <remarks/>
        public CT_LinearShadeProperties lin {
            get {
                return this.linField;
            }
            set {
                this.linField = value;
            }
        }
        
        /// <remarks/>
        public CT_PathShadeProperties path {
            get {
                return this.pathField;
            }
            set {
                this.pathField = value;
            }
        }
        
        /// <remarks/>
        public CT_RelativeRect tileRect {
            get {
                return this.tileRectField;
            }
            set {
                this.tileRectField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TileFlipMode flip {
            get {
                return this.flipField;
            }
            set {
                this.flipField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool flipSpecified {
            get {
                return this.flipFieldSpecified;
            }
            set {
                this.flipFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public bool rotWithShape {
            get {
                return this.rotWithShapeField;
            }
            set {
                this.rotWithShapeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool rotWithShapeSpecified {
            get {
                return this.rotWithShapeFieldSpecified;
            }
            set {
                this.rotWithShapeFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_TileInfoProperties {
        
        private long txField;
        
        private bool txFieldSpecified;
        
        private long tyField;
        
        private bool tyFieldSpecified;
        
        private int sxField;
        
        private bool sxFieldSpecified;
        
        private int syField;
        
        private bool syFieldSpecified;
        
        private ST_TileFlipMode flipField;
        
        private bool flipFieldSpecified;
        
        private ST_RectAlignment algnField;
        
        private bool algnFieldSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public long tx {
            get {
                return this.txField;
            }
            set {
                this.txField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool txSpecified {
            get {
                return this.txFieldSpecified;
            }
            set {
                this.txFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public long ty {
            get {
                return this.tyField;
            }
            set {
                this.tyField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool tySpecified {
            get {
                return this.tyFieldSpecified;
            }
            set {
                this.tyFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int sx {
            get {
                return this.sxField;
            }
            set {
                this.sxField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool sxSpecified {
            get {
                return this.sxFieldSpecified;
            }
            set {
                this.sxFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public int sy {
            get {
                return this.syField;
            }
            set {
                this.syField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool sySpecified {
            get {
                return this.syFieldSpecified;
            }
            set {
                this.syFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_TileFlipMode flip {
            get {
                return this.flipField;
            }
            set {
                this.flipField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool flipSpecified {
            get {
                return this.flipFieldSpecified;
            }
            set {
                this.flipFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_RectAlignment algn {
            get {
                return this.algnField;
            }
            set {
                this.algnField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool algnSpecified {
            get {
                return this.algnFieldSpecified;
            }
            set {
                this.algnFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_StretchInfoProperties {
        
        private CT_RelativeRect fillRectField;

        public CT_RelativeRect AddNewFillRect()
        {
            this.fillRectField=new CT_RelativeRect();
            return this.fillRectField;
        }
        /// <remarks/>
        public CT_RelativeRect fillRect {
            get {
                return this.fillRectField;
            }
            set {
                this.fillRectField = value;
            }
        }
    }
    
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_AlphaModulateEffect {
        
        private CT_EffectContainer contField;
        
        /// <remarks/>
        public CT_EffectContainer cont {
            get {
                return this.contField;
            }
            set {
                this.contField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_EffectContainer {
        
        private object[] itemsField;
        
        private ST_EffectContainerType typeField;
        
        private string nameField;
        
        public CT_EffectContainer() {
            this.typeField = ST_EffectContainerType.sib;
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElement("alphaBiLevel", typeof(CT_AlphaBiLevelEffect))]
        [System.Xml.Serialization.XmlElement("alphaCeiling", typeof(CT_AlphaCeilingEffect))]
        [System.Xml.Serialization.XmlElement("alphaFloor", typeof(CT_AlphaFloorEffect))]
        [System.Xml.Serialization.XmlElement("alphaInv", typeof(CT_AlphaInverseEffect))]
        [System.Xml.Serialization.XmlElement("alphaMod", typeof(CT_AlphaModulateEffect))]
        [System.Xml.Serialization.XmlElement("alphaModFix", typeof(CT_AlphaModulateFixedEffect))]
        [System.Xml.Serialization.XmlElement("alphaOutset", typeof(CT_AlphaOutsetEffect))]
        [System.Xml.Serialization.XmlElement("alphaRepl", typeof(CT_AlphaReplaceEffect))]
        [System.Xml.Serialization.XmlElement("biLevel", typeof(CT_BiLevelEffect))]
        [System.Xml.Serialization.XmlElement("blend", typeof(CT_BlendEffect))]
        [System.Xml.Serialization.XmlElement("blur", typeof(CT_BlurEffect))]
        [System.Xml.Serialization.XmlElement("clrChange", typeof(CT_ColorChangeEffect))]
        [System.Xml.Serialization.XmlElement("clrRepl", typeof(CT_ColorReplaceEffect))]
        [System.Xml.Serialization.XmlElement("cont", typeof(CT_EffectContainer))]
        [System.Xml.Serialization.XmlElement("duotone", typeof(CT_DuotoneEffect))]
        [System.Xml.Serialization.XmlElement("effect", typeof(CT_EffectReference))]
        [System.Xml.Serialization.XmlElement("fill", typeof(CT_FillEffect))]
        [System.Xml.Serialization.XmlElement("fillOverlay", typeof(CT_FillOverlayEffect))]
        [System.Xml.Serialization.XmlElement("glow", typeof(CT_GlowEffect))]
        [System.Xml.Serialization.XmlElement("grayscl", typeof(CT_GrayscaleEffect))]
        [System.Xml.Serialization.XmlElement("hsl", typeof(CT_HSLEffect))]
        [System.Xml.Serialization.XmlElement("innerShdw", typeof(CT_InnerShadowEffect))]
        [System.Xml.Serialization.XmlElement("lum", typeof(CT_LuminanceEffect))]
        [System.Xml.Serialization.XmlElement("outerShdw", typeof(CT_OuterShadowEffect))]
        [System.Xml.Serialization.XmlElement("prstShdw", typeof(CT_PresetShadowEffect))]
        [System.Xml.Serialization.XmlElement("reflection", typeof(CT_ReflectionEffect))]
        [System.Xml.Serialization.XmlElement("relOff", typeof(CT_RelativeOffsetEffect))]
        [System.Xml.Serialization.XmlElement("softEdge", typeof(CT_SoftEdgesEffect))]
        [System.Xml.Serialization.XmlElement("tint", typeof(CT_TintEffect))]
        [System.Xml.Serialization.XmlElement("xfrm", typeof(CT_TransformEffect))]
        public object[] Items {
            get {
                return this.itemsField;
            }
            set {
                this.itemsField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        [DefaultValue(ST_EffectContainerType.sib)]
        public ST_EffectContainerType type {
            get {
                return this.typeField;
            }
            set {
                this.typeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string name {
            get {
                return this.nameField;
            }
            set {
                this.nameField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_BlendEffect {
        
        private CT_EffectContainer contField;
        
        private ST_BlendMode blendField;
        
        /// <remarks/>
        public CT_EffectContainer cont {
            get {
                return this.contField;
            }
            set {
                this.contField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_BlendMode blend {
            get {
                return this.blendField;
            }
            set {
                this.blendField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_BlendMode {
        
        /// <remarks/>
        over,
        
        /// <remarks/>
        mult,
        
        /// <remarks/>
        screen,
        
        /// <remarks/>
        darken,
        
        /// <remarks/>
        lighten,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_EffectReference {
        
        private string refField;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="token")]
        public string @ref {
            get {
                return this.refField;
            }
            set {
                this.refField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_FillEffect {
        
        private CT_NoFillProperties noFillField;
        
        private CT_SolidColorFillProperties solidFillField;
        
        private CT_GradientFillProperties gradFillField;
        
        private CT_BlipFillProperties blipFillField;
        
        private CT_PatternFillProperties pattFillField;
        
        private CT_GroupFillProperties grpFillField;
        
        /// <remarks/>
        public CT_NoFillProperties noFill {
            get {
                return this.noFillField;
            }
            set {
                this.noFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_SolidColorFillProperties solidFill {
            get {
                return this.solidFillField;
            }
            set {
                this.solidFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_GradientFillProperties gradFill {
            get {
                return this.gradFillField;
            }
            set {
                this.gradFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_BlipFillProperties blipFill {
            get {
                return this.blipFillField;
            }
            set {
                this.blipFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_PatternFillProperties pattFill {
            get {
                return this.pattFillField;
            }
            set {
                this.pattFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_GroupFillProperties grpFill {
            get {
                return this.grpFillField;
            }
            set {
                this.grpFillField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_BlipFillProperties {
        
        private CT_Blip blipField;
        
        private CT_RelativeRect srcRectField;
        
        private CT_TileInfoProperties tileField;
        
        private CT_StretchInfoProperties stretchField;
        
        private uint dpiField;
        
        private bool dpiFieldSpecified;
        
        private bool rotWithShapeField;
        
        private bool rotWithShapeFieldSpecified;
        
        public CT_Blip AddNewBlip()
        {
            this.blipField = new CT_Blip();
            return blipField;
        }
        public CT_StretchInfoProperties AddNewStretch()
        {
            this.stretchField = new CT_StretchInfoProperties();
            return stretchField;
        }
        /// <remarks/>
        public CT_Blip blip {
            get {
                return this.blipField;
            }
            set {
                this.blipField = value;
            }
        }
        
        /// <remarks/>
        public CT_RelativeRect srcRect {
            get {
                return this.srcRectField;
            }
            set {
                this.srcRectField = value;
            }
        }
        
        /// <remarks/>
        public CT_TileInfoProperties tile {
            get {
                return this.tileField;
            }
            set {
                this.tileField = value;
            }
        }
        
        /// <remarks/>
        public CT_StretchInfoProperties stretch {
            get {
                return this.stretchField;
            }
            set {
                this.stretchField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public uint dpi {
            get {
                return this.dpiField;
            }
            set {
                this.dpiField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool dpiSpecified {
            get {
                return this.dpiFieldSpecified;
            }
            set {
                this.dpiFieldSpecified = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public bool rotWithShape {
            get {
                return this.rotWithShapeField;
            }
            set {
                this.rotWithShapeField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool rotWithShapeSpecified {
            get {
                return this.rotWithShapeFieldSpecified;
            }
            set {
                this.rotWithShapeFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_PatternFillProperties {
        
        private CT_Color fgClrField;
        
        private CT_Color bgClrField;
        
        private ST_PresetPatternVal prstField;
        
        private bool prstFieldSpecified;
        
        /// <remarks/>
        public CT_Color fgClr {
            get {
                return this.fgClrField;
            }
            set {
                this.fgClrField = value;
            }
        }
        
        /// <remarks/>
        public CT_Color bgClr {
            get {
                return this.bgClrField;
            }
            set {
                this.bgClrField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_PresetPatternVal prst {
            get {
                return this.prstField;
            }
            set {
                this.prstField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnore]
        public bool prstSpecified {
            get {
                return this.prstFieldSpecified;
            }
            set {
                this.prstFieldSpecified = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_PresetPatternVal {
        
        /// <remarks/>
        pct5,
        
        /// <remarks/>
        pct10,
        
        /// <remarks/>
        pct20,
        
        /// <remarks/>
        pct25,
        
        /// <remarks/>
        pct30,
        
        /// <remarks/>
        pct40,
        
        /// <remarks/>
        pct50,
        
        /// <remarks/>
        pct60,
        
        /// <remarks/>
        pct70,
        
        /// <remarks/>
        pct75,
        
        /// <remarks/>
        pct80,
        
        /// <remarks/>
        pct90,
        
        /// <remarks/>
        horz,
        
        /// <remarks/>
        vert,
        
        /// <remarks/>
        ltHorz,
        
        /// <remarks/>
        ltVert,
        
        /// <remarks/>
        dkHorz,
        
        /// <remarks/>
        dkVert,
        
        /// <remarks/>
        narHorz,
        
        /// <remarks/>
        narVert,
        
        /// <remarks/>
        dashHorz,
        
        /// <remarks/>
        dashVert,
        
        /// <remarks/>
        cross,
        
        /// <remarks/>
        dnDiag,
        
        /// <remarks/>
        upDiag,
        
        /// <remarks/>
        ltDnDiag,
        
        /// <remarks/>
        ltUpDiag,
        
        /// <remarks/>
        dkDnDiag,
        
        /// <remarks/>
        dkUpDiag,
        
        /// <remarks/>
        wdDnDiag,
        
        /// <remarks/>
        wdUpDiag,
        
        /// <remarks/>
        dashDnDiag,
        
        /// <remarks/>
        dashUpDiag,
        
        /// <remarks/>
        diagCross,
        
        /// <remarks/>
        smCheck,
        
        /// <remarks/>
        lgCheck,
        
        /// <remarks/>
        smGrid,
        
        /// <remarks/>
        lgGrid,
        
        /// <remarks/>
        dotGrid,
        
        /// <remarks/>
        smConfetti,
        
        /// <remarks/>
        lgConfetti,
        
        /// <remarks/>
        horzBrick,
        
        /// <remarks/>
        diagBrick,
        
        /// <remarks/>
        solidDmnd,
        
        /// <remarks/>
        openDmnd,
        
        /// <remarks/>
        dotDmnd,
        
        /// <remarks/>
        plaid,
        
        /// <remarks/>
        sphere,
        
        /// <remarks/>
        weave,
        
        /// <remarks/>
        divot,
        
        /// <remarks/>
        shingle,
        
        /// <remarks/>
        wave,
        
        /// <remarks/>
        trellis,
        
        /// <remarks/>
        zigZag,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_GroupFillProperties {
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_FillOverlayEffect {
        
        private CT_NoFillProperties noFillField;
        
        private CT_SolidColorFillProperties solidFillField;
        
        private CT_GradientFillProperties gradFillField;
        
        private CT_BlipFillProperties blipFillField;
        
        private CT_PatternFillProperties pattFillField;
        
        private CT_GroupFillProperties grpFillField;
        
        private ST_BlendMode blendField;
        
        /// <remarks/>
        public CT_NoFillProperties noFill {
            get {
                return this.noFillField;
            }
            set {
                this.noFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_SolidColorFillProperties solidFill {
            get {
                return this.solidFillField;
            }
            set {
                this.solidFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_GradientFillProperties gradFill {
            get {
                return this.gradFillField;
            }
            set {
                this.gradFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_BlipFillProperties blipFill {
            get {
                return this.blipFillField;
            }
            set {
                this.blipFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_PatternFillProperties pattFill {
            get {
                return this.pattFillField;
            }
            set {
                this.pattFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_GroupFillProperties grpFill {
            get {
                return this.grpFillField;
            }
            set {
                this.grpFillField = value;
            }
        }
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttribute]
        public ST_BlendMode blend {
            get {
                return this.blendField;
            }
            set {
                this.blendField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=false)]
    public enum ST_EffectContainerType {
        
        /// <remarks/>
        sib,
        
        /// <remarks/>
        tree,
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_FillProperties {
        
        private CT_NoFillProperties noFillField;
        
        private CT_SolidColorFillProperties solidFillField;
        
        private CT_GradientFillProperties gradFillField;
        
        private CT_BlipFillProperties blipFillField;
        
        private CT_PatternFillProperties pattFillField;
        
        private CT_GroupFillProperties grpFillField;
        
        /// <remarks/>
        public CT_NoFillProperties noFill {
            get {
                return this.noFillField;
            }
            set {
                this.noFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_SolidColorFillProperties solidFill {
            get {
                return this.solidFillField;
            }
            set {
                this.solidFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_GradientFillProperties gradFill {
            get {
                return this.gradFillField;
            }
            set {
                this.gradFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_BlipFillProperties blipFill {
            get {
                return this.blipFillField;
            }
            set {
                this.blipFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_PatternFillProperties pattFill {
            get {
                return this.pattFillField;
            }
            set {
                this.pattFillField = value;
            }
        }
        
        /// <remarks/>
        public CT_GroupFillProperties grpFill {
            get {
                return this.grpFillField;
            }
            set {
                this.grpFillField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_EffectList {
        
        private CT_BlurEffect blurField;
        
        private CT_FillOverlayEffect fillOverlayField;
        
        private CT_GlowEffect glowField;
        
        private CT_InnerShadowEffect innerShdwField;
        
        private CT_OuterShadowEffect outerShdwField;
        
        private CT_PresetShadowEffect prstShdwField;
        
        private CT_ReflectionEffect reflectionField;
        
        private CT_SoftEdgesEffect softEdgeField;
        
        /// <remarks/>
        public CT_BlurEffect blur {
            get {
                return this.blurField;
            }
            set {
                this.blurField = value;
            }
        }
        
        /// <remarks/>
        public CT_FillOverlayEffect fillOverlay {
            get {
                return this.fillOverlayField;
            }
            set {
                this.fillOverlayField = value;
            }
        }
        
        /// <remarks/>
        public CT_GlowEffect glow {
            get {
                return this.glowField;
            }
            set {
                this.glowField = value;
            }
        }
        
        /// <remarks/>
        public CT_InnerShadowEffect innerShdw {
            get {
                return this.innerShdwField;
            }
            set {
                this.innerShdwField = value;
            }
        }
        
        /// <remarks/>
        public CT_OuterShadowEffect outerShdw {
            get {
                return this.outerShdwField;
            }
            set {
                this.outerShdwField = value;
            }
        }
        
        /// <remarks/>
        public CT_PresetShadowEffect prstShdw {
            get {
                return this.prstShdwField;
            }
            set {
                this.prstShdwField = value;
            }
        }
        
        /// <remarks/>
        public CT_ReflectionEffect reflection {
            get {
                return this.reflectionField;
            }
            set {
                this.reflectionField = value;
            }
        }
        
        /// <remarks/>
        public CT_SoftEdgesEffect softEdge {
            get {
                return this.softEdgeField;
            }
            set {
                this.softEdgeField = value;
            }
        }
    }
    
    /// <remarks/>
    [System.Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [System.Xml.Serialization.XmlType(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace="http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable=true)]
    public partial class CT_EffectProperties {
        
        private CT_EffectList effectLstField;
        
        private CT_EffectContainer effectDagField;
        
        /// <remarks/>
        public CT_EffectList effectLst {
            get {
                return this.effectLstField;
            }
            set {
                this.effectLstField = value;
            }
        }
        
        /// <remarks/>
        public CT_EffectContainer effectDag {
            get {
                return this.effectDagField;
            }
            set {
                this.effectDagField = value;
            }
        }
    }
}
