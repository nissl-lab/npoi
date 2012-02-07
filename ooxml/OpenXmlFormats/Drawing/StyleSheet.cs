using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml
{
    public class CT_BaseStylesOverride
    {

        private CT_ColorScheme clrSchemeField;

        private CT_FontScheme fontSchemeField;

        private CT_StyleMatrix fmtSchemeField;

        public CT_BaseStylesOverride()
        {
            this.fmtSchemeField = new CT_StyleMatrix();
            this.fontSchemeField = new CT_FontScheme();
            this.clrSchemeField = new CT_ColorScheme();
        }

        public CT_ColorScheme clrScheme
        {
            get
            {
                return this.clrSchemeField;
            }
            set
            {
                this.clrSchemeField = value;
            }
        }

        public CT_FontScheme fontScheme
        {
            get
            {
                return this.fontSchemeField;
            }
            set
            {
                this.fontSchemeField = value;
            }
        }

        public CT_StyleMatrix fmtScheme
        {
            get
            {
                return this.fmtSchemeField;
            }
            set
            {
                this.fmtSchemeField = value;
            }
        }
    }

    public class CT_EmptyElement
    {
    }

    public class CT_ColorMappingOverride
    {

        private object itemField;

        public object Item
        {
            get
            {
                return this.itemField;
            }
            set
            {
                this.itemField = value;
            }
        }
    }
    public class CT_ColorSchemeAndMapping
    {

        private CT_ColorScheme clrSchemeField;

        private CT_ColorMapping clrMapField;

        public CT_ColorSchemeAndMapping()
        {
            this.clrMapField = new CT_ColorMapping();
            this.clrSchemeField = new CT_ColorScheme();
        }

        public CT_ColorScheme clrScheme
        {
            get
            {
                return this.clrSchemeField;
            }
            set
            {
                this.clrSchemeField = value;
            }
        }

        public CT_ColorMapping clrMap
        {
            get
            {
                return this.clrMapField;
            }
            set
            {
                this.clrMapField = value;
            }
        }
    }
    public class CT_ColorSchemeList
    {

        private List<CT_ColorSchemeAndMapping> extraClrSchemeField;

        public CT_ColorSchemeList()
        {
            this.extraClrSchemeField = new List<CT_ColorSchemeAndMapping>();
        }

        public List<CT_ColorSchemeAndMapping> extraClrScheme
        {
            get
            {
                return this.extraClrSchemeField;
            }
            set
            {
                this.extraClrSchemeField = value;
            }
        }
    }
    public enum ST_ColorSchemeIndex
    {

        /// <remarks/>
        dk1,

        /// <remarks/>
        lt1,

        /// <remarks/>
        dk2,

        /// <remarks/>
        lt2,

        /// <remarks/>
        accent1,

        /// <remarks/>
        accent2,

        /// <remarks/>
        accent3,

        /// <remarks/>
        accent4,

        /// <remarks/>
        accent5,

        /// <remarks/>
        accent6,

        /// <remarks/>
        hlink,

        /// <remarks/>
        folHlink,
    }
    public class CT_ColorMapping
    {

        private CT_OfficeArtExtensionList extLstField;

        private ST_ColorSchemeIndex bg1Field;

        private ST_ColorSchemeIndex tx1Field;

        private ST_ColorSchemeIndex bg2Field;

        private ST_ColorSchemeIndex tx2Field;

        private ST_ColorSchemeIndex accent1Field;

        private ST_ColorSchemeIndex accent2Field;

        private ST_ColorSchemeIndex accent3Field;

        private ST_ColorSchemeIndex accent4Field;

        private ST_ColorSchemeIndex accent5Field;

        private ST_ColorSchemeIndex accent6Field;

        private ST_ColorSchemeIndex hlinkField;

        private ST_ColorSchemeIndex folHlinkField;

        public CT_ColorMapping()
        {
            this.extLstField = new CT_OfficeArtExtensionList();
        }

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

        public ST_ColorSchemeIndex bg1
        {
            get
            {
                return this.bg1Field;
            }
            set
            {
                this.bg1Field = value;
            }
        }

        public ST_ColorSchemeIndex tx1
        {
            get
            {
                return this.tx1Field;
            }
            set
            {
                this.tx1Field = value;
            }
        }

        public ST_ColorSchemeIndex bg2
        {
            get
            {
                return this.bg2Field;
            }
            set
            {
                this.bg2Field = value;
            }
        }

        public ST_ColorSchemeIndex tx2
        {
            get
            {
                return this.tx2Field;
            }
            set
            {
                this.tx2Field = value;
            }
        }

        public ST_ColorSchemeIndex accent1
        {
            get
            {
                return this.accent1Field;
            }
            set
            {
                this.accent1Field = value;
            }
        }

        public ST_ColorSchemeIndex accent2
        {
            get
            {
                return this.accent2Field;
            }
            set
            {
                this.accent2Field = value;
            }
        }

        public ST_ColorSchemeIndex accent3
        {
            get
            {
                return this.accent3Field;
            }
            set
            {
                this.accent3Field = value;
            }
        }

        public ST_ColorSchemeIndex accent4
        {
            get
            {
                return this.accent4Field;
            }
            set
            {
                this.accent4Field = value;
            }
        }

        public ST_ColorSchemeIndex accent5
        {
            get
            {
                return this.accent5Field;
            }
            set
            {
                this.accent5Field = value;
            }
        }

        public ST_ColorSchemeIndex accent6
        {
            get
            {
                return this.accent6Field;
            }
            set
            {
                this.accent6Field = value;
            }
        }

        public ST_ColorSchemeIndex hlink
        {
            get
            {
                return this.hlinkField;
            }
            set
            {
                this.hlinkField = value;
            }
        }

        public ST_ColorSchemeIndex folHlink
        {
            get
            {
                return this.folHlinkField;
            }
            set
            {
                this.folHlinkField = value;
            }
        }
    }
    public partial class CT_ClipboardStyleSheet
    {

        private CT_BaseStyles themeElementsField;

        private CT_ColorMapping clrMapField;

        public CT_ClipboardStyleSheet()
        {
            this.clrMapField = new CT_ColorMapping();
            this.themeElementsField = new CT_BaseStyles();
        }

        public CT_BaseStyles themeElements
        {
            get
            {
                return this.themeElementsField;
            }
            set
            {
                this.themeElementsField = value;
            }
        }

        public CT_ColorMapping clrMap
        {
            get
            {
                return this.clrMapField;
            }
            set
            {
                this.clrMapField = value;
            }
        }
    }
    public partial class CT_DefaultShapeDefinition
    {

        private CT_ShapeProperties spPrField;

        private CT_TextBodyProperties bodyPrField;

        private CT_TextListStyle lstStyleField;

        private CT_ShapeStyle styleField;

        private CT_OfficeArtExtensionList extLstField;

        public CT_DefaultShapeDefinition()
        {
            this.extLstField = new CT_OfficeArtExtensionList();
            this.styleField = new CT_ShapeStyle();
            this.lstStyleField = new CT_TextListStyle();
            this.bodyPrField = new CT_TextBodyProperties();
            this.spPrField = new CT_ShapeProperties();
        }

        public CT_ShapeProperties spPr
        {
            get
            {
                return this.spPrField;
            }
            set
            {
                this.spPrField = value;
            }
        }

        public CT_TextBodyProperties bodyPr
        {
            get
            {
                return this.bodyPrField;
            }
            set
            {
                this.bodyPrField = value;
            }
        }

        public CT_TextListStyle lstStyle
        {
            get
            {
                return this.lstStyleField;
            }
            set
            {
                this.lstStyleField = value;
            }
        }

        public CT_ShapeStyle style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }

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
    }
    public partial class CT_ObjectStyleDefaults
    {

        private CT_DefaultShapeDefinition spDefField;

        private CT_DefaultShapeDefinition lnDefField;

        private CT_DefaultShapeDefinition txDefField;

        private CT_OfficeArtExtensionList extLstField;

        public CT_ObjectStyleDefaults()
        {
            this.extLstField = new CT_OfficeArtExtensionList();
            this.txDefField = new CT_DefaultShapeDefinition();
            this.lnDefField = new CT_DefaultShapeDefinition();
            this.spDefField = new CT_DefaultShapeDefinition();
        }

        public CT_DefaultShapeDefinition spDef
        {
            get
            {
                return this.spDefField;
            }
            set
            {
                this.spDefField = value;
            }
        }

        public CT_DefaultShapeDefinition lnDef
        {
            get
            {
                return this.lnDefField;
            }
            set
            {
                this.lnDefField = value;
            }
        }

        public CT_DefaultShapeDefinition txDef
        {
            get
            {
                return this.txDefField;
            }
            set
            {
                this.txDefField = value;
            }
        }

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
    }
    [XmlRoot(ElementName="theme")]
    public partial class CT_OfficeStyleSheet
    {

        private CT_BaseStyles themeElementsField;

        private CT_ObjectStyleDefaults objectDefaultsField;

        private List<CT_ColorSchemeAndMapping> extraClrSchemeLstField;

        private List<CT_CustomColor> custClrLstField;

        private CT_OfficeArtExtensionList extLstField;

        private string nameField;

        public CT_OfficeStyleSheet()
        {
            this.extLstField = new CT_OfficeArtExtensionList();
            this.custClrLstField = new List<CT_CustomColor>();
            this.extraClrSchemeLstField = new List<CT_ColorSchemeAndMapping>();
            this.objectDefaultsField = new CT_ObjectStyleDefaults();
            this.themeElementsField = new CT_BaseStyles();
            this.nameField = "";
        }

        public CT_BaseStyles themeElements
        {
            get
            {
                return this.themeElementsField;
            }
            set
            {
                this.themeElementsField = value;
            }
        }

        public CT_ObjectStyleDefaults objectDefaults
        {
            get
            {
                return this.objectDefaultsField;
            }
            set
            {
                this.objectDefaultsField = value;
            }
        }

        [System.Xml.Serialization.XmlArrayAttribute(Order = 2)]
        [System.Xml.Serialization.XmlArrayItemAttribute("extraClrScheme", IsNullable = false)]
        public List<CT_ColorSchemeAndMapping> extraClrSchemeLst
        {
            get
            {
                return this.extraClrSchemeLstField;
            }
            set
            {
                this.extraClrSchemeLstField = value;
            }
        }

        [System.Xml.Serialization.XmlArrayAttribute(Order = 3)]
        [System.Xml.Serialization.XmlArrayItemAttribute("custClr", IsNullable = false)]
        public List<CT_CustomColor> custClrLst
        {
            get
            {
                return this.custClrLstField;
            }
            set
            {
                this.custClrLstField = value;
            }
        }

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

        [System.ComponentModel.DefaultValueAttribute("")]
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

}
