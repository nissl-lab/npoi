using System;
using System.ComponentModel;
using System.Xml.Serialization;
using System.Collections.Generic;

namespace NPOI.OpenXmlFormats.Dml
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot("blip", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = false)]
    public class CT_Blip
    {

        private List<object> itemsField;

        private CT_OfficeArtExtensionList extLstField;

        private string embedField;

        private string linkField;

        private ST_BlipCompression cstateField;

        public CT_Blip()
        {
            //this.extLstField = new CT_OfficeArtExtensionList();
            //this.itemsField = new List<object>();
            this.embedField = "";
            this.linkField = "";
            this.cstateField = ST_BlipCompression.none;
        }

        [XmlElement("alphaBiLevel", typeof(CT_AlphaBiLevelEffect), Order = 0)]
        [XmlElement("alphaCeiling", typeof(CT_AlphaCeilingEffect), Order = 0)]
        [XmlElement("alphaFloor", typeof(CT_AlphaFloorEffect), Order = 0)]
        [XmlElement("alphaInv", typeof(CT_AlphaInverseEffect), Order = 0)]
        [XmlElement("alphaMod", typeof(CT_AlphaModulateEffect), Order = 0)]
        [XmlElement("alphaModFix", typeof(CT_AlphaModulateFixedEffect), Order = 0)]
        [XmlElement("alphaRepl", typeof(CT_AlphaReplaceEffect), Order = 0)]
        [XmlElement("biLevel", typeof(CT_BiLevelEffect), Order = 0)]
        [XmlElement("blur", typeof(CT_BlurEffect), Order = 0)]
        [XmlElement("clrChange", typeof(CT_ColorChangeEffect), Order = 0)]
        [XmlElement("clrRepl", typeof(CT_ColorReplaceEffect), Order = 0)]
        [XmlElement("duotone", typeof(CT_DuotoneEffect), Order = 0)]
        [XmlElement("fillOverlay", typeof(CT_FillOverlayEffect), Order = 0)]
        [XmlElement("grayscl", typeof(CT_GrayscaleEffect), Order = 0)]
        [XmlElement("hsl", typeof(CT_HSLEffect), Order = 0)]
        [XmlElement("lum", typeof(CT_LuminanceEffect), Order = 0)]
        [XmlElement("tint", typeof(CT_TintEffect), Order = 0)]
        public List<object> Items
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

        [XmlElement(Order = 1)]
        public CT_OfficeArtExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        [DefaultValue("")]
        public string embed
        {
            get
            {
                return this.embedField;
            }
            set
            {
                this.embedField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        [DefaultValue("")]
        public string link
        {
            get
            {
                return this.linkField;
            }
            set
            {
                this.linkField = value;
            }
        }

        [XmlAttribute]
        [DefaultValue(ST_BlipCompression.none)]
        public ST_BlipCompression cstate
        {
            get
            {
                return this.cstateField;
            }
            set
            {
                this.cstateField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AlphaBiLevelEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_TransformEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_TintEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_SoftEdgesEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_RelativeOffsetEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_ReflectionEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_PresetShadowEffect
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
            this.prstClrField = new CT_PresetColor();
            this.schemeClrField = new CT_SchemeColor();
            this.sysClrField = new CT_SystemColor();
            this.hslClrField = new CT_HslColor();
            this.srgbClrField = new CT_SRgbColor();
            this.scrgbClrField = new CT_ScRgbColor();
            this.distField = ((long)(0));
            this.dirField = 0;
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_PresetShadowVal
    {

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

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_OuterShadowEffect
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
            this.prstClrField = new CT_PresetColor();
            this.schemeClrField = new CT_SchemeColor();
            this.sysClrField = new CT_SystemColor();
            this.hslClrField = new CT_HslColor();
            this.srgbClrField = new CT_SRgbColor();
            this.scrgbClrField = new CT_ScRgbColor();
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

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_LuminanceEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_InnerShadowEffect
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
            this.prstClrField = new CT_PresetColor();
            this.schemeClrField = new CT_SchemeColor();
            this.sysClrField = new CT_SystemColor();
            this.hslClrField = new CT_HslColor();
            this.srgbClrField = new CT_SRgbColor();
            this.scrgbClrField = new CT_ScRgbColor();
            this.blurRadField = ((long)(0));
            this.distField = ((long)(0));
            this.dirField = 0;
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_HSLEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GrayscaleEffect
    {
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GlowEffect
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
            this.prstClrField = new CT_PresetColor();
            this.schemeClrField = new CT_SchemeColor();
            this.sysClrField = new CT_SystemColor();
            this.hslClrField = new CT_HslColor();
            this.srgbClrField = new CT_SRgbColor();
            this.scrgbClrField = new CT_ScRgbColor();
            this.radField = ((long)(0));
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_FillOverlayEffect
    {

        private CT_NoFillProperties noFillField;

        private CT_SolidColorFillProperties solidFillField;

        private CT_GradientFillProperties gradFillField;

        private CT_BlipFillProperties blipFillField;

        private CT_PatternFillProperties pattFillField;

        private CT_GroupFillProperties grpFillField;

        private ST_BlendMode blendField;

        public CT_FillOverlayEffect()
        {
            //this.grpFillField = new CT_GroupFillProperties();
            //this.pattFillField = new CT_PatternFillProperties();
            //this.blipFillField = new CT_BlipFillProperties();
            //this.gradFillField = new CT_GradientFillProperties();
            //this.solidFillField = new CT_SolidColorFillProperties();
            //this.noFillField = new CT_NoFillProperties();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_NoFillProperties
    {
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_SolidColorFillProperties
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        public CT_SolidColorFillProperties()
        {
            //this.prstClrField = new CT_PresetColor();
            //this.schemeClrField = new CT_SchemeColor();
            //this.sysClrField = new CT_SystemColor();
            //this.hslClrField = new CT_HslColor();
            //this.srgbClrField = new CT_SRgbColor();
            //this.scrgbClrField = new CT_ScRgbColor();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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

        public bool IsSetSrgbClr()
        {
            throw new NotImplementedException();
        }

        public CT_SRgbColor AddNewSrgbClr()
        {
            throw new NotImplementedException();
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GradientFillProperties
    {

        private List<CT_GradientStop> gsLstField;

        private CT_LinearShadeProperties linField;

        private CT_PathShadeProperties pathField;

        private CT_RelativeRect tileRectField;

        private ST_TileFlipMode flipField;

        private bool flipFieldSpecified;

        private bool rotWithShapeField;

        private bool rotWithShapeFieldSpecified;

        public CT_GradientFillProperties()
        {
            //this.tileRectField = new CT_RelativeRect();
            //this.pathField = new CT_PathShadeProperties();
            //this.linField = new CT_LinearShadeProperties();
            //this.gsLstField = new List<CT_GradientStop>();
        }

        [XmlArray(Order = 0)]
        [XmlArrayItem("gs", IsNullable = false)]
        public List<CT_GradientStop> gsLst
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GradientStop
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        private int posField;

        public CT_GradientStop()
        {
            //this.prstClrField = new CT_PresetColor();
            //this.schemeClrField = new CT_SchemeColor();
            //this.sysClrField = new CT_SystemColor();
            //this.hslClrField = new CT_HslColor();
            //this.srgbClrField = new CT_SRgbColor();
            //this.scrgbClrField = new CT_ScRgbColor();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_LinearShadeProperties
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
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_PathShadeProperties
    {

        private CT_RelativeRect fillToRectField;

        private ST_PathShadeType pathField;

        private bool pathFieldSpecified;

        public CT_PathShadeProperties()
        {
            this.fillToRectField = new CT_RelativeRect();
        }

        [XmlElement(Order = 0)]
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
    public enum ST_PathShadeType
    {

        /// <remarks/>
        shape,

        /// <remarks/>
        circle,

        /// <remarks/>
        rect,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_TileFlipMode
    {

        /// <remarks/>
        none,

        /// <remarks/>
        x,

        /// <remarks/>
        y,

        /// <remarks/>
        xy,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_BlipFillProperties
    {

        private CT_Blip blipField = null;

        private CT_RelativeRect srcRectField = null;

        private CT_TileInfoProperties tileField = null;

        private CT_StretchInfoProperties stretchField = null;

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

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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
            get 
            { 
                return (uint)this.dpiField; 
            }
            set 
            { 
                this.dpiField = value; 
            }
        }
        [XmlIgnore]
        public bool dpiSpecified
        {
            get
            {
                return dpiFieldSpecified;
            }
            set
            {
                this.dpiFieldSpecified = value;
            }
        }


        [XmlAttribute]
        public bool rotWithShape
        {
            get
            {
                return (bool)this.rotWithShapeField;
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
                return rotWithShapeFieldSpecified;
            }
            set
            {
                this.rotWithShapeFieldSpecified = value;
            }
        }

        public bool IsSetBlip()
        {
            return this.blipField != null;
        }
    }



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_TileInfoProperties
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_StretchInfoProperties
    {

        private CT_RelativeRect fillRectField = null;

        public CT_RelativeRect AddNewFillRect()
        {
            this.fillRectField = new CT_RelativeRect();
            return this.fillRectField;
        }

        [XmlElement(Order = 0)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_PatternFillProperties
    {

        private CT_Color fgClrField;

        private CT_Color bgClrField;

        private ST_PresetPatternVal prstField;

        private bool prstFieldSpecified;

        public CT_PatternFillProperties()
        {
            this.bgClrField = new CT_Color();
            this.fgClrField = new CT_Color();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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
    public enum ST_PresetPatternVal
    {

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


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GroupFillProperties
    {
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_BlendMode
    {

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

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_FillEffect
    {

        private CT_NoFillProperties noFillField;

        private CT_SolidColorFillProperties solidFillField;

        private CT_GradientFillProperties gradFillField;

        private CT_BlipFillProperties blipFillField;

        private CT_PatternFillProperties pattFillField;

        private CT_GroupFillProperties grpFillField;

        public CT_FillEffect()
        {
            this.grpFillField = new CT_GroupFillProperties();
            this.pattFillField = new CT_PatternFillProperties();
            this.blipFillField = new CT_BlipFillProperties();
            this.gradFillField = new CT_GradientFillProperties();
            this.solidFillField = new CT_SolidColorFillProperties();
            this.noFillField = new CT_NoFillProperties();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_DuotoneEffect
    {
        private List<object> itemsField;

        public CT_DuotoneEffect()
        {
            this.itemsField = new List<object>();
        }

        [XmlElement("hslClr", typeof(CT_HslColor), Order = 0)]
        [XmlElement("prstClr", typeof(CT_PresetColor), Order = 0)]
        [XmlElement("schemeClr", typeof(CT_SchemeColor), Order = 0)]
        [XmlElement("scrgbClr", typeof(CT_ScRgbColor), Order = 0)]
        [XmlElement("srgbClr", typeof(CT_SRgbColor), Order = 0)]
        [XmlElement("sysClr", typeof(CT_SystemColor), Order = 0)]
        public List<object> Items
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

        #region another way
        //private CT_ScRgbColor[] scrgbClrField;

        //private CT_SRgbColor[] srgbClrField;

        //private CT_HslColor[] hslClrField;

        //private CT_SystemColor[] sysClrField;

        //private CT_SchemeColor[] schemeClrField;

        //private CT_PresetColor[] prstClrField;


        //[XmlElement("scrgbClr")]
        //public CT_ScRgbColor[] scrgbClr
        //{
        //    get
        //    {
        //        return this.scrgbClrField;
        //    }
        //    set
        //    {
        //        this.scrgbClrField = value;
        //    }
        //}


        //[XmlElement("srgbClr")]
        //public CT_SRgbColor[] srgbClr
        //{
        //    get
        //    {
        //        return this.srgbClrField;
        //    }
        //    set
        //    {
        //        this.srgbClrField = value;
        //    }
        //}


        //[XmlElement("hslClr")]
        //public CT_HslColor[] hslClr
        //{
        //    get
        //    {
        //        return this.hslClrField;
        //    }
        //    set
        //    {
        //        this.hslClrField = value;
        //    }
        //}


        //[XmlElement("sysClr")]
        //public CT_SystemColor[] sysClr
        //{
        //    get
        //    {
        //        return this.sysClrField;
        //    }
        //    set
        //    {
        //        this.sysClrField = value;
        //    }
        //}


        //[XmlElement("schemeClr")]
        //public CT_SchemeColor[] schemeClr
        //{
        //    get
        //    {
        //        return this.schemeClrField;
        //    }
        //    set
        //    {
        //        this.schemeClrField = value;
        //    }
        //}


        //[XmlElement("prstClr")]
        //public CT_PresetColor[] prstClr
        //{
        //    get
        //    {
        //        return this.prstClrField;
        //    }
        //    set
        //    {
        //        this.prstClrField = value;
        //    }
        //}

        #endregion
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_ColorReplaceEffect
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        public CT_ColorReplaceEffect()
        {
            this.prstClrField = new CT_PresetColor();
            this.schemeClrField = new CT_SchemeColor();
            this.sysClrField = new CT_SystemColor();
            this.hslClrField = new CT_HslColor();
            this.srgbClrField = new CT_SRgbColor();
            this.scrgbClrField = new CT_ScRgbColor();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_ColorChangeEffect
    {

        private CT_Color clrFromField;

        private CT_Color clrToField;

        private bool useAField;

        public CT_ColorChangeEffect()
        {
            this.clrToField = new CT_Color();
            this.clrFromField = new CT_Color();
            this.useAField = true;
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_BlurEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_BlendEffect
    {

        private CT_EffectContainer contField;

        private ST_BlendMode blendField;

        public CT_BlendEffect()
        {
            this.contField = new CT_EffectContainer();
        }

        [XmlElement(Order = 0)]
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
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_EffectContainer
    {

        private List<object> itemsField;

        private ST_EffectContainerType typeField;

        private string nameField;

        public CT_EffectContainer()
        {
            this.itemsField = new List<object>();
            this.typeField = ST_EffectContainerType.sib;
        }

        [XmlElement("alphaBiLevel", typeof(CT_AlphaBiLevelEffect), Order = 0)]
        [XmlElement("alphaCeiling", typeof(CT_AlphaCeilingEffect), Order = 0)]
        [XmlElement("alphaFloor", typeof(CT_AlphaFloorEffect), Order = 0)]
        [XmlElement("alphaInv", typeof(CT_AlphaInverseEffect), Order = 0)]
        [XmlElement("alphaMod", typeof(CT_AlphaModulateEffect), Order = 0)]
        [XmlElement("alphaModFix", typeof(CT_AlphaModulateFixedEffect), Order = 0)]
        [XmlElement("alphaOutset", typeof(CT_AlphaOutsetEffect), Order = 0)]
        [XmlElement("alphaRepl", typeof(CT_AlphaReplaceEffect), Order = 0)]
        [XmlElement("biLevel", typeof(CT_BiLevelEffect), Order = 0)]
        [XmlElement("blend", typeof(CT_BlendEffect), Order = 0)]
        [XmlElement("blur", typeof(CT_BlurEffect), Order = 0)]
        [XmlElement("clrChange", typeof(CT_ColorChangeEffect), Order = 0)]
        [XmlElement("clrRepl", typeof(CT_ColorReplaceEffect), Order = 0)]
        [XmlElement("cont", typeof(CT_EffectContainer), Order = 0)]
        [XmlElement("duotone", typeof(CT_DuotoneEffect), Order = 0)]
        [XmlElement("effect", typeof(CT_EffectReference), Order = 0)]
        [XmlElement("fill", typeof(CT_FillEffect), Order = 0)]
        [XmlElement("fillOverlay", typeof(CT_FillOverlayEffect), Order = 0)]
        [XmlElement("glow", typeof(CT_GlowEffect), Order = 0)]
        [XmlElement("grayscl", typeof(CT_GrayscaleEffect), Order = 0)]
        [XmlElement("hsl", typeof(CT_HSLEffect), Order = 0)]
        [XmlElement("innerShdw", typeof(CT_InnerShadowEffect), Order = 0)]
        [XmlElement("lum", typeof(CT_LuminanceEffect), Order = 0)]
        [XmlElement("outerShdw", typeof(CT_OuterShadowEffect), Order = 0)]
        [XmlElement("prstShdw", typeof(CT_PresetShadowEffect), Order = 0)]
        [XmlElement("reflection", typeof(CT_ReflectionEffect), Order = 0)]
        [XmlElement("relOff", typeof(CT_RelativeOffsetEffect), Order = 0)]
        [XmlElement("softEdge", typeof(CT_SoftEdgesEffect), Order = 0)]
        [XmlElement("tint", typeof(CT_TintEffect), Order = 0)]
        [XmlElement("xfrm", typeof(CT_TransformEffect), Order = 0)]
        public List<object> Items
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AlphaCeilingEffect
    {
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AlphaFloorEffect
    {
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AlphaInverseEffect
    {

        private CT_ScRgbColor scrgbClrField;

        private CT_SRgbColor srgbClrField;

        private CT_HslColor hslClrField;

        private CT_SystemColor sysClrField;

        private CT_SchemeColor schemeClrField;

        private CT_PresetColor prstClrField;

        public CT_AlphaInverseEffect()
        {
            this.prstClrField = new CT_PresetColor();
            this.schemeClrField = new CT_SchemeColor();
            this.sysClrField = new CT_SystemColor();
            this.hslClrField = new CT_HslColor();
            this.srgbClrField = new CT_SRgbColor();
            this.scrgbClrField = new CT_ScRgbColor();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AlphaModulateEffect
    {

        private CT_EffectContainer contField;

        public CT_AlphaModulateEffect()
        {
            this.contField = new CT_EffectContainer();
        }

        [XmlElement(Order = 0)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AlphaModulateFixedEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AlphaOutsetEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_AlphaReplaceEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_BiLevelEffect
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_EffectReference
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_EffectContainerType
    {

        /// <remarks/>
        sib,

        /// <remarks/>
        tree,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public enum ST_BlipCompression
    {

        /// <remarks/>
        email,

        /// <remarks/>
        screen,

        /// <remarks/>
        print,

        /// <remarks/>
        hqprint,

        /// <remarks/>
        none,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_GradientStopList
    {

        private List<CT_GradientStop> gsField;

        public CT_GradientStopList()
        {
            this.gsField = new List<CT_GradientStop>();
        }

        [XmlElement("gs", Order = 0)]
        public List<CT_GradientStop> gs
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_FillProperties
    {

        private CT_NoFillProperties noFillField;

        private CT_SolidColorFillProperties solidFillField;

        private CT_GradientFillProperties gradFillField;

        private CT_BlipFillProperties blipFillField;

        private CT_PatternFillProperties pattFillField;

        private CT_GroupFillProperties grpFillField;

        public CT_FillProperties()
        {
            this.grpFillField = new CT_GroupFillProperties();
            this.pattFillField = new CT_PatternFillProperties();
            this.blipFillField = new CT_BlipFillProperties();
            this.gradFillField = new CT_GradientFillProperties();
            this.solidFillField = new CT_SolidColorFillProperties();
            this.noFillField = new CT_NoFillProperties();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_EffectList
    {

        private CT_BlurEffect blurField;

        private CT_FillOverlayEffect fillOverlayField;

        private CT_GlowEffect glowField;

        private CT_InnerShadowEffect innerShdwField;

        private CT_OuterShadowEffect outerShdwField;

        private CT_PresetShadowEffect prstShdwField;

        private CT_ReflectionEffect reflectionField;

        private CT_SoftEdgesEffect softEdgeField;

        public CT_EffectList()
        {
            this.softEdgeField = new CT_SoftEdgesEffect();
            this.reflectionField = new CT_ReflectionEffect();
            this.prstShdwField = new CT_PresetShadowEffect();
            this.outerShdwField = new CT_OuterShadowEffect();
            this.innerShdwField = new CT_InnerShadowEffect();
            this.glowField = new CT_GlowEffect();
            this.fillOverlayField = new CT_FillOverlayEffect();
            this.blurField = new CT_BlurEffect();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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

        [XmlElement(Order = 2)]
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

        [XmlElement(Order = 3)]
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

        [XmlElement(Order = 4)]
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

        [XmlElement(Order = 5)]
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

        [XmlElement(Order = 6)]
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

        [XmlElement(Order = 7)]
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public class CT_EffectProperties
    {

        private CT_EffectList effectLstField;

        private CT_EffectContainer effectDagField;

        public CT_EffectProperties()
        {
            this.effectDagField = new CT_EffectContainer();
            this.effectLstField = new CT_EffectList();
        }

        [XmlElement(Order = 0)]
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

        [XmlElement(Order = 1)]
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
