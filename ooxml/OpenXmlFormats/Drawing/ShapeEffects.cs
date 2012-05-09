using System;
using System.ComponentModel;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml
{



    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_AlphaBiLevelEffect
    {

        private int threshField;


        [XmlAttribute]
        public int thresh
        {
            get
            {
                return this.threshField;
            }
            set
            {
                this.threshField = value;
            }
        }
    }



    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_HslColor
    {

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


        [XmlElement("tint")]
        public CT_PositiveFixedPercentage[] tint
        {
            get
            {
                return this.tintField;
            }
            set
            {
                this.tintField = value;
            }
        }


        [XmlElement("shade")]
        public CT_PositiveFixedPercentage[] shade
        {
            get
            {
                return this.shadeField;
            }
            set
            {
                this.shadeField = value;
            }
        }


        [XmlElement("comp")]
        public CT_ComplementTransform[] comp
        {
            get
            {
                return this.compField;
            }
            set
            {
                this.compField = value;
            }
        }


        [XmlElement("inv")]
        public CT_InverseTransform[] inv
        {
            get
            {
                return this.invField;
            }
            set
            {
                this.invField = value;
            }
        }


        [XmlElement("gray")]
        public CT_GrayscaleTransform[] gray
        {
            get
            {
                return this.grayField;
            }
            set
            {
                this.grayField = value;
            }
        }


        [XmlElement("alpha")]
        public CT_PositiveFixedPercentage[] alpha
        {
            get
            {
                return this.alphaField;
            }
            set
            {
                this.alphaField = value;
            }
        }


        [XmlElement("alphaOff")]
        public CT_FixedPercentage[] alphaOff
        {
            get
            {
                return this.alphaOffField;
            }
            set
            {
                this.alphaOffField = value;
            }
        }


        [XmlElement("alphaMod")]
        public CT_PositivePercentage[] alphaMod
        {
            get
            {
                return this.alphaModField;
            }
            set
            {
                this.alphaModField = value;
            }
        }


        [XmlElement("hue")]
        public CT_PositiveFixedAngle[] hue
        {
            get
            {
                return this.hueField;
            }
            set
            {
                this.hueField = value;
            }
        }


        [XmlElement("hueOff")]
        public CT_Angle[] hueOff
        {
            get
            {
                return this.hueOffField;
            }
            set
            {
                this.hueOffField = value;
            }
        }


        [XmlElement("hueMod")]
        public CT_PositivePercentage[] hueMod
        {
            get
            {
                return this.hueModField;
            }
            set
            {
                this.hueModField = value;
            }
        }


        [XmlElement("sat")]
        public CT_Percentage[] sat
        {
            get
            {
                return this.satField;
            }
            set
            {
                this.satField = value;
            }
        }


        [XmlElement("satOff")]
        public CT_Percentage[] satOff
        {
            get
            {
                return this.satOffField;
            }
            set
            {
                this.satOffField = value;
            }
        }


        [XmlElement("satMod")]
        public CT_Percentage[] satMod
        {
            get
            {
                return this.satModField;
            }
            set
            {
                this.satModField = value;
            }
        }


        [XmlElement("lum")]
        public CT_Percentage[] lum
        {
            get
            {
                return this.lumField;
            }
            set
            {
                this.lumField = value;
            }
        }


        [XmlElement("lumOff")]
        public CT_Percentage[] lumOff
        {
            get
            {
                return this.lumOffField;
            }
            set
            {
                this.lumOffField = value;
            }
        }


        [XmlElement("lumMod")]
        public CT_Percentage[] lumMod
        {
            get
            {
                return this.lumModField;
            }
            set
            {
                this.lumModField = value;
            }
        }


        [XmlElement("red")]
        public CT_Percentage[] red
        {
            get
            {
                return this.redField;
            }
            set
            {
                this.redField = value;
            }
        }


        [XmlElement("redOff")]
        public CT_Percentage[] redOff
        {
            get
            {
                return this.redOffField;
            }
            set
            {
                this.redOffField = value;
            }
        }


        [XmlElement("redMod")]
        public CT_Percentage[] redMod
        {
            get
            {
                return this.redModField;
            }
            set
            {
                this.redModField = value;
            }
        }


        [XmlElement("green")]
        public CT_Percentage[] green
        {
            get
            {
                return this.greenField;
            }
            set
            {
                this.greenField = value;
            }
        }


        [XmlElement("greenOff")]
        public CT_Percentage[] greenOff
        {
            get
            {
                return this.greenOffField;
            }
            set
            {
                this.greenOffField = value;
            }
        }


        [XmlElement("greenMod")]
        public CT_Percentage[] greenMod
        {
            get
            {
                return this.greenModField;
            }
            set
            {
                this.greenModField = value;
            }
        }


        [XmlElement("blue")]
        public CT_Percentage[] blue
        {
            get
            {
                return this.blueField;
            }
            set
            {
                this.blueField = value;
            }
        }


        [XmlElement("blueOff")]
        public CT_Percentage[] blueOff
        {
            get
            {
                return this.blueOffField;
            }
            set
            {
                this.blueOffField = value;
            }
        }


        [XmlElement("blueMod")]
        public CT_Percentage[] blueMod
        {
            get
            {
                return this.blueModField;
            }
            set
            {
                this.blueModField = value;
            }
        }


        [XmlElement("gamma")]
        public CT_GammaTransform[] gamma
        {
            get
            {
                return this.gammaField;
            }
            set
            {
                this.gammaField = value;
            }
        }


        [XmlElement("invGamma")]
        public CT_InverseGammaTransform[] invGamma
        {
            get
            {
                return this.invGammaField;
            }
            set
            {
                this.invGammaField = value;
            }
        }


        [XmlAttribute("hue")]
        public int hue1
        {
            get
            {
                return this.hue1Field;
            }
            set
            {
                this.hue1Field = value;
            }
        }


        [XmlAttribute("sat")]
        public int sat1
        {
            get
            {
                return this.sat1Field;
            }
            set
            {
                this.sat1Field = value;
            }
        }


        [XmlAttribute("lum")]
        public int lum1
        {
            get
            {
                return this.lum1Field;
            }
            set
            {
                this.lum1Field = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_SystemColor
    {

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


        [XmlElement("tint")]
        public CT_PositiveFixedPercentage[] tint
        {
            get
            {
                return this.tintField;
            }
            set
            {
                this.tintField = value;
            }
        }


        [XmlElement("shade")]
        public CT_PositiveFixedPercentage[] shade
        {
            get
            {
                return this.shadeField;
            }
            set
            {
                this.shadeField = value;
            }
        }


        [XmlElement("comp")]
        public CT_ComplementTransform[] comp
        {
            get
            {
                return this.compField;
            }
            set
            {
                this.compField = value;
            }
        }


        [XmlElement("inv")]
        public CT_InverseTransform[] inv
        {
            get
            {
                return this.invField;
            }
            set
            {
                this.invField = value;
            }
        }


        [XmlElement("gray")]
        public CT_GrayscaleTransform[] gray
        {
            get
            {
                return this.grayField;
            }
            set
            {
                this.grayField = value;
            }
        }


        [XmlElement("alpha")]
        public CT_PositiveFixedPercentage[] alpha
        {
            get
            {
                return this.alphaField;
            }
            set
            {
                this.alphaField = value;
            }
        }


        [XmlElement("alphaOff")]
        public CT_FixedPercentage[] alphaOff
        {
            get
            {
                return this.alphaOffField;
            }
            set
            {
                this.alphaOffField = value;
            }
        }


        [XmlElement("alphaMod")]
        public CT_PositivePercentage[] alphaMod
        {
            get
            {
                return this.alphaModField;
            }
            set
            {
                this.alphaModField = value;
            }
        }


        [XmlElement("hue")]
        public CT_PositiveFixedAngle[] hue
        {
            get
            {
                return this.hueField;
            }
            set
            {
                this.hueField = value;
            }
        }


        [XmlElement("hueOff")]
        public CT_Angle[] hueOff
        {
            get
            {
                return this.hueOffField;
            }
            set
            {
                this.hueOffField = value;
            }
        }


        [XmlElement("hueMod")]
        public CT_PositivePercentage[] hueMod
        {
            get
            {
                return this.hueModField;
            }
            set
            {
                this.hueModField = value;
            }
        }


        [XmlElement("sat")]
        public CT_Percentage[] sat
        {
            get
            {
                return this.satField;
            }
            set
            {
                this.satField = value;
            }
        }


        [XmlElement("satOff")]
        public CT_Percentage[] satOff
        {
            get
            {
                return this.satOffField;
            }
            set
            {
                this.satOffField = value;
            }
        }


        [XmlElement("satMod")]
        public CT_Percentage[] satMod
        {
            get
            {
                return this.satModField;
            }
            set
            {
                this.satModField = value;
            }
        }


        [XmlElement("lum")]
        public CT_Percentage[] lum
        {
            get
            {
                return this.lumField;
            }
            set
            {
                this.lumField = value;
            }
        }


        [XmlElement("lumOff")]
        public CT_Percentage[] lumOff
        {
            get
            {
                return this.lumOffField;
            }
            set
            {
                this.lumOffField = value;
            }
        }


        [XmlElement("lumMod")]
        public CT_Percentage[] lumMod
        {
            get
            {
                return this.lumModField;
            }
            set
            {
                this.lumModField = value;
            }
        }


        [XmlElement("red")]
        public CT_Percentage[] red
        {
            get
            {
                return this.redField;
            }
            set
            {
                this.redField = value;
            }
        }


        [XmlElement("redOff")]
        public CT_Percentage[] redOff
        {
            get
            {
                return this.redOffField;
            }
            set
            {
                this.redOffField = value;
            }
        }


        [XmlElement("redMod")]
        public CT_Percentage[] redMod
        {
            get
            {
                return this.redModField;
            }
            set
            {
                this.redModField = value;
            }
        }


        [XmlElement("green")]
        public CT_Percentage[] green
        {
            get
            {
                return this.greenField;
            }
            set
            {
                this.greenField = value;
            }
        }


        [XmlElement("greenOff")]
        public CT_Percentage[] greenOff
        {
            get
            {
                return this.greenOffField;
            }
            set
            {
                this.greenOffField = value;
            }
        }


        [XmlElement("greenMod")]
        public CT_Percentage[] greenMod
        {
            get
            {
                return this.greenModField;
            }
            set
            {
                this.greenModField = value;
            }
        }


        [XmlElement("blue")]
        public CT_Percentage[] blue
        {
            get
            {
                return this.blueField;
            }
            set
            {
                this.blueField = value;
            }
        }


        [XmlElement("blueOff")]
        public CT_Percentage[] blueOff
        {
            get
            {
                return this.blueOffField;
            }
            set
            {
                this.blueOffField = value;
            }
        }


        [XmlElement("blueMod")]
        public CT_Percentage[] blueMod
        {
            get
            {
                return this.blueModField;
            }
            set
            {
                this.blueModField = value;
            }
        }


        [XmlElement("gamma")]
        public CT_GammaTransform[] gamma
        {
            get
            {
                return this.gammaField;
            }
            set
            {
                this.gammaField = value;
            }
        }


        [XmlElement("invGamma")]
        public CT_InverseGammaTransform[] invGamma
        {
            get
            {
                return this.invGammaField;
            }
            set
            {
                this.invGammaField = value;
            }
        }


        [XmlAttribute]
        public ST_SystemColorVal val
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


        [XmlAttribute(DataType = "hexBinary")]
        public byte[] lastClr
        {
            get
            {
                return this.lastClrField;
            }
            set
            {
                this.lastClrField = value;
            }
        }
    }




    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_AlphaCeilingEffect
    {
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_AlphaFloorEffect
    {
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_AlphaInverseEffect
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;


        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }


        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }


        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }


        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }


        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }


        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_AlphaModulateFixedEffect
    {

        private int amtField;

        public CT_AlphaModulateFixedEffect()
        {
            this.amtField = 100000;
        }


        [XmlAttribute]
        [DefaultValue(100000)]
        public int amt
        {
            get
            {
                return this.amtField;
            }
            set
            {
                this.amtField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_AlphaOutsetEffect
    {

        private long radField;

        public CT_AlphaOutsetEffect()
        {
            this.radField = ((long)(0));
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long rad
        {
            get
            {
                return this.radField;
            }
            set
            {
                this.radField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_AlphaReplaceEffect
    {

        private int aField;


        [XmlAttribute]
        public int a
        {
            get
            {
                return this.aField;
            }
            set
            {
                this.aField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_BiLevelEffect
    {

        private int threshField;


        [XmlAttribute]
        public int thresh
        {
            get
            {
                return this.threshField;
            }
            set
            {
                this.threshField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_BlurEffect
    {

        private long radField;

        private bool growField;

        public CT_BlurEffect()
        {
            this.radField = ((long)(0));
            this.growField = true;
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long rad
        {
            get
            {
                return this.radField;
            }
            set
            {
                this.radField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(true)]
        public bool grow
        {
            get
            {
                return this.growField;
            }
            set
            {
                this.growField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_ColorChangeEffect
    {

        private CT_Color clrFromField;

        private CT_Color clrToField;

        private bool useAField;

        public CT_ColorChangeEffect()
        {
            this.useAField = true;
        }


        public CT_Color clrFrom
        {
            get
            {
                return this.clrFromField;
            }
            set
            {
                this.clrFromField = value;
            }
        }


        public CT_Color clrTo
        {
            get
            {
                return this.clrToField;
            }
            set
            {
                this.clrToField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(true)]
        public bool useA
        {
            get
            {
                return this.useAField;
            }
            set
            {
                this.useAField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_ColorReplaceEffect
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;


        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }


        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }


        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }


        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }


        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }


        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_DuotoneEffect
    {

        private CT_ScRgbColor[] scrgbClrField;

        private CT_SRgbColor[] srgbClrField;

        private CT_HslColor[] hslClrField;

        private CT_SystemColor[] sysClrField;

        private CT_SchemeColor[] schemeClrField;

        private CT_PresetColor[] prstClrField;


        [XmlElement("scrgbClr")]
        public CT_ScRgbColor[] scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }


        [XmlElement("srgbClr")]
        public CT_SRgbColor[] srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }


        [XmlElement("hslClr")]
        public CT_HslColor[] hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }


        [XmlElement("sysClr")]
        public CT_SystemColor[] sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }


        [XmlElement("schemeClr")]
        public CT_SchemeColor[] schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }


        [XmlElement("prstClr")]
        public CT_PresetColor[] prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_GlowEffect
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        private long radField;

        public CT_GlowEffect()
        {
            this.radField = ((long)(0));
        }


        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }


        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }


        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }


        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }


        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }


        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long rad
        {
            get
            {
                return this.radField;
            }
            set
            {
                this.radField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_GrayscaleEffect
    {
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_HSLEffect
    {

        private int hueField;

        private int satField;

        private int lumField;

        public CT_HSLEffect()
        {
            this.hueField = 0;
            this.satField = 0;
            this.lumField = 0;
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int hue
        {
            get
            {
                return this.hueField;
            }
            set
            {
                this.hueField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int sat
        {
            get
            {
                return this.satField;
            }
            set
            {
                this.satField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int lum
        {
            get
            {
                return this.lumField;
            }
            set
            {
                this.lumField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_InnerShadowEffect
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        private long blurRadField;

        private long distField;

        private int dirField;

        public CT_InnerShadowEffect()
        {
            this.blurRadField = ((long)(0));
            this.distField = ((long)(0));
            this.dirField = 0;
        }


        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }


        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }


        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }


        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }


        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }


        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long blurRad
        {
            get
            {
                return this.blurRadField;
            }
            set
            {
                this.blurRadField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long dist
        {
            get
            {
                return this.distField;
            }
            set
            {
                this.distField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int dir
        {
            get
            {
                return this.dirField;
            }
            set
            {
                this.dirField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_LuminanceEffect
    {

        private int brightField;

        private int contrastField;

        public CT_LuminanceEffect()
        {
            this.brightField = 0;
            this.contrastField = 0;
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int bright
        {
            get
            {
                return this.brightField;
            }
            set
            {
                this.brightField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int contrast
        {
            get
            {
                return this.contrastField;
            }
            set
            {
                this.contrastField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_OuterShadowEffect
    {

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

        public CT_OuterShadowEffect()
        {
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


        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }


        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }


        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }


        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }


        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }


        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long blurRad
        {
            get
            {
                return this.blurRadField;
            }
            set
            {
                this.blurRadField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long dist
        {
            get
            {
                return this.distField;
            }
            set
            {
                this.distField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int dir
        {
            get
            {
                return this.dirField;
            }
            set
            {
                this.dirField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(100000)]
        public int sx
        {
            get
            {
                return this.sxField;
            }
            set
            {
                this.sxField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(100000)]
        public int sy
        {
            get
            {
                return this.syField;
            }
            set
            {
                this.syField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int kx
        {
            get
            {
                return this.kxField;
            }
            set
            {
                this.kxField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int ky
        {
            get
            {
                return this.kyField;
            }
            set
            {
                this.kyField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(ST_RectAlignment.b)]
        public ST_RectAlignment algn
        {
            get
            {
                return this.algnField;
            }
            set
            {
                this.algnField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(true)]
        public bool rotWithShape
        {
            get
            {
                return this.rotWithShapeField;
            }
            set
            {
                this.rotWithShapeField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = false)]
    public enum ST_PresetShadowVal
    {


        shdw1,


        shdw2,


        shdw3,


        shdw4,


        shdw5,


        shdw6,


        shdw7,


        shdw8,


        shdw9,


        shdw10,


        shdw11,


        shdw12,


        shdw13,


        shdw14,


        shdw15,


        shdw16,


        shdw17,


        shdw18,


        shdw19,


        shdw20,
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_PresetShadowEffect
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        private ST_PresetShadowVal prstField;

        private long distField;

        private int dirField;

        public CT_PresetShadowEffect()
        {
            this.distField = ((long)(0));
            this.dirField = 0;
        }


        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }


        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }


        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }


        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }


        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }


        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }


        [XmlAttribute]
        public ST_PresetShadowVal prst
        {
            get
            {
                return this.prstField;
            }
            set
            {
                this.prstField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long dist
        {
            get
            {
                return this.distField;
            }
            set
            {
                this.distField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int dir
        {
            get
            {
                return this.dirField;
            }
            set
            {
                this.dirField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_ReflectionEffect
    {

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

        public CT_ReflectionEffect()
        {
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


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long blurRad
        {
            get
            {
                return this.blurRadField;
            }
            set
            {
                this.blurRadField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(100000)]
        public int stA
        {
            get
            {
                return this.stAField;
            }
            set
            {
                this.stAField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int stPos
        {
            get
            {
                return this.stPosField;
            }
            set
            {
                this.stPosField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int endA
        {
            get
            {
                return this.endAField;
            }
            set
            {
                this.endAField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(100000)]
        public int endPos
        {
            get
            {
                return this.endPosField;
            }
            set
            {
                this.endPosField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long dist
        {
            get
            {
                return this.distField;
            }
            set
            {
                this.distField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int dir
        {
            get
            {
                return this.dirField;
            }
            set
            {
                this.dirField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(5400000)]
        public int fadeDir
        {
            get
            {
                return this.fadeDirField;
            }
            set
            {
                this.fadeDirField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(100000)]
        public int sx
        {
            get
            {
                return this.sxField;
            }
            set
            {
                this.sxField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(100000)]
        public int sy
        {
            get
            {
                return this.syField;
            }
            set
            {
                this.syField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int kx
        {
            get
            {
                return this.kxField;
            }
            set
            {
                this.kxField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int ky
        {
            get
            {
                return this.kyField;
            }
            set
            {
                this.kyField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(ST_RectAlignment.b)]
        public ST_RectAlignment algn
        {
            get
            {
                return this.algnField;
            }
            set
            {
                this.algnField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(true)]
        public bool rotWithShape
        {
            get
            {
                return this.rotWithShapeField;
            }
            set
            {
                this.rotWithShapeField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_RelativeOffsetEffect
    {

        private int txField;

        private int tyField;

        public CT_RelativeOffsetEffect()
        {
            this.txField = 0;
            this.tyField = 0;
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int tx
        {
            get
            {
                return this.txField;
            }
            set
            {
                this.txField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int ty
        {
            get
            {
                return this.tyField;
            }
            set
            {
                this.tyField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_SoftEdgesEffect
    {

        private long radField;


        [XmlAttribute]
        public long rad
        {
            get
            {
                return this.radField;
            }
            set
            {
                this.radField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_TintEffect
    {

        private int hueField;

        private int amtField;

        public CT_TintEffect()
        {
            this.hueField = 0;
            this.amtField = 0;
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int hue
        {
            get
            {
                return this.hueField;
            }
            set
            {
                this.hueField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int amt
        {
            get
            {
                return this.amtField;
            }
            set
            {
                this.amtField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_TransformEffect
    {

        private int sxField;

        private int syField;

        private int kxField;

        private int kyField;

        private long txField;

        private long tyField;

        public CT_TransformEffect()
        {
            this.sxField = 100000;
            this.syField = 100000;
            this.kxField = 0;
            this.kyField = 0;
            this.txField = ((long)(0));
            this.tyField = ((long)(0));
        }


        [XmlAttribute]
        [DefaultValue(100000)]
        public int sx
        {
            get
            {
                return this.sxField;
            }
            set
            {
                this.sxField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(100000)]
        public int sy
        {
            get
            {
                return this.syField;
            }
            set
            {
                this.syField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int kx
        {
            get
            {
                return this.kxField;
            }
            set
            {
                this.kxField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(0)]
        public int ky
        {
            get
            {
                return this.kyField;
            }
            set
            {
                this.kyField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long tx
        {
            get
            {
                return this.txField;
            }
            set
            {
                this.txField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(typeof(long), "0")]
        public long ty
        {
            get
            {
                return this.tyField;
            }
            set
            {
                this.tyField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public class CT_NoFillProperties
    {
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_SolidColorFillProperties
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;


        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }


        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }


        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }


        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }


        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }


        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_LinearShadeProperties
    {

        private int angField;

        private bool angFieldSpecified;

        private bool scaledField;

        private bool scaledFieldSpecified;


        [XmlAttribute]
        public int ang
        {
            get
            {
                return this.angField;
            }
            set
            {
                this.angField = value;
            }
        }


        [XmlIgnore]
        public bool angSpecified
        {
            get
            {
                return this.angFieldSpecified;
            }
            set
            {
                this.angFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool scaled
        {
            get
            {
                return this.scaledField;
            }
            set
            {
                this.scaledField = value;
            }
        }


        [XmlIgnore]
        public bool scaledSpecified
        {
            get
            {
                return this.scaledFieldSpecified;
            }
            set
            {
                this.scaledFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = false)]
    public enum ST_PathShadeType
    {


        shape,


        circle,


        rect,
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_PathShadeProperties
    {

        private CT_RelativeRect fillToRectField;

        private ST_PathShadeType pathField;

        private bool pathFieldSpecified;


        public CT_RelativeRect fillToRect
        {
            get
            {
                return this.fillToRectField;
            }
            set
            {
                this.fillToRectField = value;
            }
        }


        [XmlAttribute]
        public ST_PathShadeType path
        {
            get
            {
                return this.pathField;
            }
            set
            {
                this.pathField = value;
            }
        }


        [XmlIgnore]
        public bool pathSpecified
        {
            get
            {
                return this.pathFieldSpecified;
            }
            set
            {
                this.pathFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = false)]
    public enum ST_TileFlipMode
    {


        none,


        x,


        y,


        xy,
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_GradientStop
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        private int posField;


        public CT_ScRgbColor scrgbClr
        {
            get
            {
                return this.scrgbClrField;
            }
            set
            {
                this.scrgbClrField = value;
            }
        }


        public CT_SRgbColor srgbClr
        {
            get
            {
                return this.srgbClrField;
            }
            set
            {
                this.srgbClrField = value;
            }
        }


        public CT_HslColor hslClr
        {
            get
            {
                return this.hslClrField;
            }
            set
            {
                this.hslClrField = value;
            }
        }


        public CT_SystemColor sysClr
        {
            get
            {
                return this.sysClrField;
            }
            set
            {
                this.sysClrField = value;
            }
        }


        public CT_SchemeColor schemeClr
        {
            get
            {
                return this.schemeClrField;
            }
            set
            {
                this.schemeClrField = value;
            }
        }


        public CT_PresetColor prstClr
        {
            get
            {
                return this.prstClrField;
            }
            set
            {
                this.prstClrField = value;
            }
        }


        [XmlAttribute]
        public int pos
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
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_GradientStopList
    {

        private CT_GradientStop[] gsField;


        [XmlElement("gs")]
        public CT_GradientStop[] gs
        {
            get
            {
                return this.gsField;
            }
            set
            {
                this.gsField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_GradientFillProperties
    {

        private CT_GradientStop[] gsLstField;

        private CT_LinearShadeProperties linField;

        private CT_PathShadeProperties pathField;

        private CT_RelativeRect tileRectField;

        private ST_TileFlipMode flipField;

        private bool flipFieldSpecified;

        private bool rotWithShapeField;

        private bool rotWithShapeFieldSpecified;


        [XmlArrayItem("gs", IsNullable = false)]
        public CT_GradientStop[] gsLst
        {
            get
            {
                return this.gsLstField;
            }
            set
            {
                this.gsLstField = value;
            }
        }


        public CT_LinearShadeProperties lin
        {
            get
            {
                return this.linField;
            }
            set
            {
                this.linField = value;
            }
        }


        public CT_PathShadeProperties path
        {
            get
            {
                return this.pathField;
            }
            set
            {
                this.pathField = value;
            }
        }


        public CT_RelativeRect tileRect
        {
            get
            {
                return this.tileRectField;
            }
            set
            {
                this.tileRectField = value;
            }
        }


        [XmlAttribute]
        public ST_TileFlipMode flip
        {
            get
            {
                return this.flipField;
            }
            set
            {
                this.flipField = value;
            }
        }


        [XmlIgnore]
        public bool flipSpecified
        {
            get
            {
                return this.flipFieldSpecified;
            }
            set
            {
                this.flipFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool rotWithShape
        {
            get
            {
                return this.rotWithShapeField;
            }
            set
            {
                this.rotWithShapeField = value;
            }
        }


        [XmlIgnore]
        public bool rotWithShapeSpecified
        {
            get
            {
                return this.rotWithShapeFieldSpecified;
            }
            set
            {
                this.rotWithShapeFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_TileInfoProperties
    {

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


        [XmlAttribute]
        public long tx
        {
            get
            {
                return this.txField;
            }
            set
            {
                this.txField = value;
            }
        }


        [XmlIgnore]
        public bool txSpecified
        {
            get
            {
                return this.txFieldSpecified;
            }
            set
            {
                this.txFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public long ty
        {
            get
            {
                return this.tyField;
            }
            set
            {
                this.tyField = value;
            }
        }


        [XmlIgnore]
        public bool tySpecified
        {
            get
            {
                return this.tyFieldSpecified;
            }
            set
            {
                this.tyFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int sx
        {
            get
            {
                return this.sxField;
            }
            set
            {
                this.sxField = value;
            }
        }


        [XmlIgnore]
        public bool sxSpecified
        {
            get
            {
                return this.sxFieldSpecified;
            }
            set
            {
                this.sxFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public int sy
        {
            get
            {
                return this.syField;
            }
            set
            {
                this.syField = value;
            }
        }


        [XmlIgnore]
        public bool sySpecified
        {
            get
            {
                return this.syFieldSpecified;
            }
            set
            {
                this.syFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_TileFlipMode flip
        {
            get
            {
                return this.flipField;
            }
            set
            {
                this.flipField = value;
            }
        }


        [XmlIgnore]
        public bool flipSpecified
        {
            get
            {
                return this.flipFieldSpecified;
            }
            set
            {
                this.flipFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public ST_RectAlignment algn
        {
            get
            {
                return this.algnField;
            }
            set
            {
                this.algnField = value;
            }
        }


        [XmlIgnore]
        public bool algnSpecified
        {
            get
            {
                return this.algnFieldSpecified;
            }
            set
            {
                this.algnFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public class CT_StretchInfoProperties
    {

        private CT_RelativeRect fillRectField = null;

        public CT_RelativeRect AddNewFillRect()
        {
            this.fillRectField = new CT_RelativeRect();
            return this.fillRectField;
        }

        [XmlElement]
        public CT_RelativeRect fillRect
        {
            get
            {
                return this.fillRectField;
            }
            set
            {
                this.fillRectField = value;
            }
        }
    }



    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_AlphaModulateEffect
    {

        private CT_EffectContainer contField;


        public CT_EffectContainer cont
        {
            get
            {
                return this.contField;
            }
            set
            {
                this.contField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_EffectContainer
    {

        private object[] itemsField;

        private ST_EffectContainerType typeField;

        private string nameField;

        public CT_EffectContainer()
        {
            this.typeField = ST_EffectContainerType.sib;
        }


        [XmlElement("alphaBiLevel", typeof(CT_AlphaBiLevelEffect))]
        [XmlElement("alphaCeiling", typeof(CT_AlphaCeilingEffect))]
        [XmlElement("alphaFloor", typeof(CT_AlphaFloorEffect))]
        [XmlElement("alphaInv", typeof(CT_AlphaInverseEffect))]
        [XmlElement("alphaMod", typeof(CT_AlphaModulateEffect))]
        [XmlElement("alphaModFix", typeof(CT_AlphaModulateFixedEffect))]
        [XmlElement("alphaOutset", typeof(CT_AlphaOutsetEffect))]
        [XmlElement("alphaRepl", typeof(CT_AlphaReplaceEffect))]
        [XmlElement("biLevel", typeof(CT_BiLevelEffect))]
        [XmlElement("blend", typeof(CT_BlendEffect))]
        [XmlElement("blur", typeof(CT_BlurEffect))]
        [XmlElement("clrChange", typeof(CT_ColorChangeEffect))]
        [XmlElement("clrRepl", typeof(CT_ColorReplaceEffect))]
        [XmlElement("cont", typeof(CT_EffectContainer))]
        [XmlElement("duotone", typeof(CT_DuotoneEffect))]
        [XmlElement("effect", typeof(CT_EffectReference))]
        [XmlElement("fill", typeof(CT_FillEffect))]
        [XmlElement("fillOverlay", typeof(CT_FillOverlayEffect))]
        [XmlElement("glow", typeof(CT_GlowEffect))]
        [XmlElement("grayscl", typeof(CT_GrayscaleEffect))]
        [XmlElement("hsl", typeof(CT_HSLEffect))]
        [XmlElement("innerShdw", typeof(CT_InnerShadowEffect))]
        [XmlElement("lum", typeof(CT_LuminanceEffect))]
        [XmlElement("outerShdw", typeof(CT_OuterShadowEffect))]
        [XmlElement("prstShdw", typeof(CT_PresetShadowEffect))]
        [XmlElement("reflection", typeof(CT_ReflectionEffect))]
        [XmlElement("relOff", typeof(CT_RelativeOffsetEffect))]
        [XmlElement("softEdge", typeof(CT_SoftEdgesEffect))]
        [XmlElement("tint", typeof(CT_TintEffect))]
        [XmlElement("xfrm", typeof(CT_TransformEffect))]
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


        [XmlAttribute]
        [DefaultValue(ST_EffectContainerType.sib)]
        public ST_EffectContainerType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }


        [XmlAttribute(DataType = "token")]
        public string name
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
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_BlendEffect
    {

        private CT_EffectContainer contField;

        private ST_BlendMode blendField;


        public CT_EffectContainer cont
        {
            get
            {
                return this.contField;
            }
            set
            {
                this.contField = value;
            }
        }


        [XmlAttribute]
        public ST_BlendMode blend
        {
            get
            {
                return this.blendField;
            }
            set
            {
                this.blendField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = false)]
    public enum ST_BlendMode
    {


        over,


        mult,


        screen,


        darken,


        lighten,
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_EffectReference
    {

        private string refField;


        [XmlAttribute(DataType = "token")]
        public string @ref
        {
            get
            {
                return this.refField;
            }
            set
            {
                this.refField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_FillEffect
    {

        private CT_NoFillProperties noFillField;

        private CT_SolidColorFillProperties solidFillField;

        private CT_GradientFillProperties gradFillField;

        private CT_BlipFillProperties blipFillField;

        private CT_PatternFillProperties pattFillField;

        private CT_GroupFillProperties grpFillField;


        public CT_NoFillProperties noFill
        {
            get
            {
                return this.noFillField;
            }
            set
            {
                this.noFillField = value;
            }
        }


        public CT_SolidColorFillProperties solidFill
        {
            get
            {
                return this.solidFillField;
            }
            set
            {
                this.solidFillField = value;
            }
        }


        public CT_GradientFillProperties gradFill
        {
            get
            {
                return this.gradFillField;
            }
            set
            {
                this.gradFillField = value;
            }
        }


        public CT_BlipFillProperties blipFill
        {
            get
            {
                return this.blipFillField;
            }
            set
            {
                this.blipFillField = value;
            }
        }


        public CT_PatternFillProperties pattFill
        {
            get
            {
                return this.pattFillField;
            }
            set
            {
                this.pattFillField = value;
            }
        }


        public CT_GroupFillProperties grpFill
        {
            get
            {
                return this.grpFillField;
            }
            set
            {
                this.grpFillField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_BlipFillProperties
    {

        private CT_Blip blipField = null;

        private CT_RelativeRect srcRectField = null;

        private CT_TileInfoProperties tileField = null;

        private CT_StretchInfoProperties stretchField = null;

        private uint? dpiField = null;

        private bool? rotWithShapeField = null;


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

        [XmlElement]
        public CT_Blip blip
        {
            get
            {
                return this.blipField;
            }
            set
            {
                this.blipField = value;
            }
        }


        [XmlElement]
        public CT_RelativeRect srcRect
        {
            get
            {
                return this.srcRectField;
            }
            set
            {
                this.srcRectField = value;
            }
        }

        [XmlElement]
        public CT_TileInfoProperties tile
        {
            get
            {
                return this.tileField;
            }
            set
            {
                this.tileField = value;
            }
        }

        [XmlElement]
        public CT_StretchInfoProperties stretch
        {
            get
            {
                return this.stretchField;
            }
            set
            {
                this.stretchField = value;
            }
        }


        [XmlAttribute]
        public uint dpi
        {
            get { return (uint)this.dpiField; }
            set { this.dpiField = value; }
        }
        [XmlIgnore]
        public bool dpiSpecified
        {
            get { return null != this.dpiField; }
        }


        [XmlAttribute]
        public bool rotWithShape
        {
            get { return (bool)this.rotWithShapeField; }
            set { this.rotWithShapeField = value; }
        }
        [XmlIgnore]
        public bool rotWithShapeSpecified
        {
            get { return null != this.rotWithShapeField; }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_PatternFillProperties
    {

        private CT_Color fgClrField;

        private CT_Color bgClrField;

        private ST_PresetPatternVal prstField;

        private bool prstFieldSpecified;


        public CT_Color fgClr
        {
            get
            {
                return this.fgClrField;
            }
            set
            {
                this.fgClrField = value;
            }
        }


        public CT_Color bgClr
        {
            get
            {
                return this.bgClrField;
            }
            set
            {
                this.bgClrField = value;
            }
        }


        [XmlAttribute]
        public ST_PresetPatternVal prst
        {
            get
            {
                return this.prstField;
            }
            set
            {
                this.prstField = value;
            }
        }


        [XmlIgnore]
        public bool prstSpecified
        {
            get
            {
                return this.prstFieldSpecified;
            }
            set
            {
                this.prstFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = false)]
    public enum ST_PresetPatternVal
    {


        pct5,


        pct10,


        pct20,


        pct25,


        pct30,


        pct40,


        pct50,


        pct60,


        pct70,


        pct75,


        pct80,


        pct90,


        horz,


        vert,


        ltHorz,


        ltVert,


        dkHorz,


        dkVert,


        narHorz,


        narVert,


        dashHorz,


        dashVert,


        cross,


        dnDiag,


        upDiag,


        ltDnDiag,


        ltUpDiag,


        dkDnDiag,


        dkUpDiag,


        wdDnDiag,


        wdUpDiag,


        dashDnDiag,


        dashUpDiag,


        diagCross,


        smCheck,


        lgCheck,


        smGrid,


        lgGrid,


        dotGrid,


        smConfetti,


        lgConfetti,


        horzBrick,


        diagBrick,


        solidDmnd,


        openDmnd,


        dotDmnd,


        plaid,


        sphere,


        weave,


        divot,


        shingle,


        wave,


        trellis,


        zigZag,
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_GroupFillProperties
    {
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_FillOverlayEffect
    {

        private CT_NoFillProperties noFillField;

        private CT_SolidColorFillProperties solidFillField;

        private CT_GradientFillProperties gradFillField;

        private CT_BlipFillProperties blipFillField;

        private CT_PatternFillProperties pattFillField;

        private CT_GroupFillProperties grpFillField;

        private ST_BlendMode blendField;


        public CT_NoFillProperties noFill
        {
            get
            {
                return this.noFillField;
            }
            set
            {
                this.noFillField = value;
            }
        }


        public CT_SolidColorFillProperties solidFill
        {
            get
            {
                return this.solidFillField;
            }
            set
            {
                this.solidFillField = value;
            }
        }


        public CT_GradientFillProperties gradFill
        {
            get
            {
                return this.gradFillField;
            }
            set
            {
                this.gradFillField = value;
            }
        }


        public CT_BlipFillProperties blipFill
        {
            get
            {
                return this.blipFillField;
            }
            set
            {
                this.blipFillField = value;
            }
        }


        public CT_PatternFillProperties pattFill
        {
            get
            {
                return this.pattFillField;
            }
            set
            {
                this.pattFillField = value;
            }
        }


        public CT_GroupFillProperties grpFill
        {
            get
            {
                return this.grpFillField;
            }
            set
            {
                this.grpFillField = value;
            }
        }


        [XmlAttribute]
        public ST_BlendMode blend
        {
            get
            {
                return this.blendField;
            }
            set
            {
                this.blendField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = false)]
    public enum ST_EffectContainerType
    {


        sib,


        tree,
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_FillProperties
    {

        private CT_NoFillProperties noFillField;

        private CT_SolidColorFillProperties solidFillField;

        private CT_GradientFillProperties gradFillField;

        private CT_BlipFillProperties blipFillField;

        private CT_PatternFillProperties pattFillField;

        private CT_GroupFillProperties grpFillField;


        public CT_NoFillProperties noFill
        {
            get
            {
                return this.noFillField;
            }
            set
            {
                this.noFillField = value;
            }
        }


        public CT_SolidColorFillProperties solidFill
        {
            get
            {
                return this.solidFillField;
            }
            set
            {
                this.solidFillField = value;
            }
        }


        public CT_GradientFillProperties gradFill
        {
            get
            {
                return this.gradFillField;
            }
            set
            {
                this.gradFillField = value;
            }
        }


        public CT_BlipFillProperties blipFill
        {
            get
            {
                return this.blipFillField;
            }
            set
            {
                this.blipFillField = value;
            }
        }


        public CT_PatternFillProperties pattFill
        {
            get
            {
                return this.pattFillField;
            }
            set
            {
                this.pattFillField = value;
            }
        }


        public CT_GroupFillProperties grpFill
        {
            get
            {
                return this.grpFillField;
            }
            set
            {
                this.grpFillField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_EffectList
    {

        private CT_BlurEffect blurField;

        private CT_FillOverlayEffect fillOverlayField;

        private CT_GlowEffect glowField;

        private CT_InnerShadowEffect innerShdwField;

        private CT_OuterShadowEffect outerShdwField;

        private CT_PresetShadowEffect prstShdwField;

        private CT_ReflectionEffect reflectionField;

        private CT_SoftEdgesEffect softEdgeField;


        public CT_BlurEffect blur
        {
            get
            {
                return this.blurField;
            }
            set
            {
                this.blurField = value;
            }
        }


        public CT_FillOverlayEffect fillOverlay
        {
            get
            {
                return this.fillOverlayField;
            }
            set
            {
                this.fillOverlayField = value;
            }
        }


        public CT_GlowEffect glow
        {
            get
            {
                return this.glowField;
            }
            set
            {
                this.glowField = value;
            }
        }


        public CT_InnerShadowEffect innerShdw
        {
            get
            {
                return this.innerShdwField;
            }
            set
            {
                this.innerShdwField = value;
            }
        }


        public CT_OuterShadowEffect outerShdw
        {
            get
            {
                return this.outerShdwField;
            }
            set
            {
                this.outerShdwField = value;
            }
        }


        public CT_PresetShadowEffect prstShdw
        {
            get
            {
                return this.prstShdwField;
            }
            set
            {
                this.prstShdwField = value;
            }
        }


        public CT_ReflectionEffect reflection
        {
            get
            {
                return this.reflectionField;
            }
            set
            {
                this.reflectionField = value;
            }
        }


        public CT_SoftEdgesEffect softEdge
        {
            get
            {
                return this.softEdgeField;
            }
            set
            {
                this.softEdgeField = value;
            }
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_EffectProperties
    {

        private CT_EffectList effectLstField;

        private CT_EffectContainer effectDagField;


        public CT_EffectList effectLst
        {
            get
            {
                return this.effectLstField;
            }
            set
            {
                this.effectLstField = value;
            }
        }


        public CT_EffectContainer effectDag
        {
            get
            {
                return this.effectDagField;
            }
            set
            {
                this.effectDagField = value;
            }
        }
    }
}
