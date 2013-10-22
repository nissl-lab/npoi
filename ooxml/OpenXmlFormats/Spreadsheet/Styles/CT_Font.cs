using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "font")]
    public class CT_Font
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Font));
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
            new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });

        // all elements are optional
        private List<CT_FontName> nameField = null; // name of the font
        private List<CT_IntProperty> charsetField = null;
        private List<CT_IntProperty> familyField = null; // family of the font
        private List<CT_BooleanProperty> bField = null; // typeface bold
        private List<CT_BooleanProperty> iField = null;   // italic
        private List<CT_BooleanProperty> strikeField = null; //   strike through
        private CT_BooleanProperty outlineField = null;
        private CT_BooleanProperty shadowField = null;
        private CT_BooleanProperty condenseField = null;
        private CT_BooleanProperty extendField = null;
        private List<CT_Color> colorField = null;
        private List<CT_FontSize> szField = null; // size of the font
        private List<CT_UnderlineProperty> uField = null; // underline
        private List<CT_VerticalAlignFontProperty> vertAlignField = null;  // vertical alignment of the text
        private List<CT_FontScheme> schemeField = null;

        public CT_Font()
        {
            this.nameField = new List<CT_FontName>();
            this.szField = new List<CT_FontSize>();
            this.colorField = new List<CT_Color>();
            this.familyField = new List<CT_IntProperty>();
            this.charsetField = new List<CT_IntProperty>();
            this.uField = new List<CT_UnderlineProperty>();
            this.bField = new List<CT_BooleanProperty>();
            this.iField = new List<CT_BooleanProperty>();
            this.vertAlignField = new List<CT_VerticalAlignFontProperty>();
            this.schemeField = new List<CT_FontScheme>();
            this.strikeField = new List<CT_BooleanProperty>();
        }

        public static CT_Font Parse(string xml)
        {
            CT_Font result;
            using (StringReader stream = new StringReader(xml))
            {
                result = (CT_Font)serializer.Deserialize(stream);
            }
            return result;
        }

        public static void Save(Stream stream, CT_Font font)
        {
            serializer.Serialize(stream, font, namespaces);
        }

        public static string GetString(CT_Font font)
        {
            using (StringWriter writer = new StringWriter())
            {
                serializer.Serialize(writer, font, namespaces);
                return writer.ToString();
            }
        }
        #region name
        [XmlElement]
        public List<CT_FontName> name
        {
            get { return this.nameField; }
            set { this.nameField = value; }
        }
        public int sizeOfNameArray()
        {
            return this.nameField.Count;
        }
        public CT_FontName AddNewName()
        {
            CT_FontName fn = new CT_FontName();
            this.nameField.Add(fn);
            return fn;
        }
        public void SetNameArray(int index, CT_FontName value)
        {
            this.nameField[index] = value;
        }
        public CT_FontName GetNameArray(int index)
        {
            return this.nameField[index];
        }
        #endregion name

        #region charset
        [XmlElement]
        public List<CT_IntProperty> charset
        {
            get { return this.charsetField; }
            set { this.charsetField = value; }
        }
        public int sizeOfCharsetArray()
        {
            return this.charsetField.Count;
        }
        public CT_IntProperty AddNewCharset()
        {
            CT_IntProperty prop = new CT_IntProperty();
            this.charsetField.Add(prop);
            return prop;
        }
        public void SetCharsetArray(int index, CT_IntProperty value)
        {
            this.charsetField[index] = value;
        }
        public CT_IntProperty GetCharsetArray(int index)
        {
            return this.charsetField[index];
        }
        #endregion charset

        #region family
        [XmlElement]
        public List<CT_IntProperty> family
        {
            get { return this.familyField; }
            set { this.familyField = value; }
        }
        public int sizeOfFamilyArray()
        {
            return this.familyField.Count;
        }
        public CT_IntProperty AddNewFamily()
        {
            CT_IntProperty newfamily = new CT_IntProperty();
            this.familyField.Add(newfamily);
            return newfamily;
        }
        public void SetFamilyArray(int index, CT_IntProperty value)
        {
            this.familyField[index] = value;
        }
        public CT_IntProperty GetFamilyArray(int index)
        {
            return this.familyField[index];
        }
        #endregion family

        #region b
        [XmlElement]
        public List<CT_BooleanProperty> b
        {
            get { return this.bField; }
            set { this.bField = value; }
        }
        public int sizeOfBArray()
        {
            if (this.bField == null)
                return 0;
            return this.bField.Count;
        }
        public CT_BooleanProperty AddNewB()
        {
            if (this.bField == null)
                this.bField = new List<CT_BooleanProperty>();
            CT_BooleanProperty newB = new CT_BooleanProperty();
            this.bField.Add(newB);
            return newB;
        }
        public void SetBArray(int index, CT_BooleanProperty value)
        {
            this.bField[index] = value;
        }
        public void SetBArray(CT_BooleanProperty[] array)
        {
            if (array == null)
                this.bField = null;
            else
                this.bField = new List<CT_BooleanProperty>(array);
        }
        public CT_BooleanProperty GetBArray(int index)
        {
            return this.bField[index];
        }
        #endregion b

        #region i
        [XmlElement]
        public List<CT_BooleanProperty> i
        {
            get { return this.iField; }
            set { this.iField = value; }
        }
        public int sizeOfIArray()
        {
            return this.iField.Count;
        }
        public CT_BooleanProperty AddNewI()
        {
            CT_BooleanProperty newI = new CT_BooleanProperty();
            this.iField.Add(newI);
            return newI;
        }
        public void SetIArray(int index, CT_BooleanProperty value)
        {
            this.iField[index] = value;
        }
        public void SetIArray(CT_BooleanProperty[] array)
        {
            if (array == null)
                this.iField = new List<CT_BooleanProperty>();
            else
                this.iField = new List<CT_BooleanProperty>(array);
        }
        public CT_BooleanProperty GetIArray(int index)
        {
            return this.iField[index];
        }
        #endregion i

        #region strike
        [XmlElement]
        public List<CT_BooleanProperty> strike
        {
            get { return this.strikeField; }
            set { this.strikeField = value; }
        }
        public int sizeOfStrikeArray()
        {
            return this.strikeField.Count;
        }
        public CT_BooleanProperty AddNewStrike()
        {
            CT_BooleanProperty prop = new CT_BooleanProperty();
            this.strikeField.Add(prop);
            return prop;
        }
        public void SetStrikeArray(int index, CT_BooleanProperty value)
        {
            this.strikeField[index] = value;
        }
        public void SetStrikeArray(CT_BooleanProperty[] array)
        {
            this.strikeField = new List<CT_BooleanProperty>(array);
        }
        public CT_BooleanProperty GetStrikeArray(int index)
        {
            return this.strikeField[index];
        }
        #endregion strike

        #region outline
        [XmlElement]
        public CT_BooleanProperty outline
        {
            get { return this.outlineField; }
            set { this.outlineField = value; }
        }
        [XmlIgnore]
        public bool outlineSpecified
        {
            get { return (null != outlineField); }
        }
        public int sizeOfOutlineArray()
        {
            return this.outlineField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewOutline()
        {
            this.outlineField = new CT_BooleanProperty();
            return this.outlineField;
        }
        public void SetOutlineArray(CT_BooleanProperty[] array)
        {
            this.outlineField = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetOutlineArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.outlineField;
        }
        #endregion outline

        #region shadow
        [XmlElement]
        public CT_BooleanProperty shadow
        {
            get { return this.shadowField; }
            set { this.shadowField = value; }
        }
        [XmlIgnore]
        public bool shadowSpecified
        {
            get { return (null != shadowField); }
        }
        public int sizeOfShadowArray()
        {
            return this.shadowField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewShadow()
        {
            this.shadowField = new CT_BooleanProperty();
            return this.shadowField;
        }
        public CT_BooleanProperty GetShadowArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.shadowField;
        }
        #endregion shadow

        #region condense
        [XmlElement]
        public CT_BooleanProperty condense
        {
            get { return this.condenseField; }
            set { this.condenseField = value; }
        }
        [XmlIgnore]
        public bool condenseSpecified
        {
            get { return (null != condenseField); }
        }
        public int sizeOfCondenseArray()
        {
            return this.condenseField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewCondense()
        {
            this.condenseField = new CT_BooleanProperty();
            return this.condenseField;
        }
        public CT_BooleanProperty GetCondenseArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.condenseField;
        }
        #endregion condense

        #region extend
        [XmlElement]
        public CT_BooleanProperty extend
        {
            get { return this.extendField; }
            set { this.extendField = value; }
        }
        [XmlIgnore]
        public bool extendSpecified
        {
            get { return (null != extendField); }
        }
        public int sizeOfExtendArray()
        {
            return this.extendField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewExtend()
        {
            this.extendField = new CT_BooleanProperty();
            return this.extendField;
        }
        public CT_BooleanProperty GetExtendArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.extendField;
        }
        #endregion extend

        #region color
        [XmlElement]
        public List<CT_Color> color
        {
            get { return this.colorField; }
            set { this.colorField = value; }
        }
        public int sizeOfColorArray()
        {
            return this.colorField.Count;
        }
        public CT_Color GetColorArray(int index)
        {
            return this.colorField[index];
        }
        public void SetColorArray(int index, CT_Color value)
        {
            this.colorField[index] = value;
        }
        public void SetColorArray(CT_Color[] array)
        {
            this.colorField = new List<CT_Color>(array);
        }
        public CT_Color AddNewColor()
        {
            CT_Color newColor = new CT_Color();
            this.colorField.Add(newColor);
            return newColor;
        }
        #endregion color

        #region sz
        [XmlElement]
        public List<CT_FontSize> sz
        {
            get { return this.szField; }
            set { this.szField = value; }
        }
        public int sizeOfSzArray()
        {
            return this.szField.Count;
        }
        public CT_FontSize AddNewSz()
        {
            CT_FontSize newFs = new CT_FontSize();
            this.szField.Add(newFs);
            return newFs;
        }
        public void SetSzArray(int index, CT_FontSize value)
        {
            this.szField[index] = value;
        }
        public void SetSzArray(CT_FontSize[] array)
        {
            this.szField = new List<CT_FontSize>(array);
        }
        public CT_FontSize GetSzArray(int index)
        {
            return this.szField[index];
        }
        #endregion sz

        #region u
        [XmlElement]
        public List<CT_UnderlineProperty> u
        {
            get { return this.uField; }
            set { this.uField = value; }
        }
        public int sizeOfUArray()
        {
            return this.uField.Count;
        }
        public CT_UnderlineProperty AddNewU()
        {
            CT_UnderlineProperty newU = new CT_UnderlineProperty();
            this.uField.Add(newU);
            return newU;
        }
        public void SetUArray(int index, CT_UnderlineProperty value)
        {
            if (uField[index] != null)
                this.uField[index] = value;
            else
                this.uField.Insert(index, value);
        }
        public void SetUArray(CT_UnderlineProperty[] array)
        {
            this.uField = new List<CT_UnderlineProperty>(array);
        }
        public CT_UnderlineProperty GetUArray(int index)
        {
            return this.uField[index];
        }
        #endregion u

        #region vertAlign
        [XmlElement]
        public List<CT_VerticalAlignFontProperty> vertAlign
        {
            get { return this.vertAlignField; }
            set { this.vertAlignField = value; }
        }
        public int sizeOfVertAlignArray()
        {
            return this.vertAlignField.Count;
        }
        public CT_VerticalAlignFontProperty AddNewVertAlign()
        {
            CT_VerticalAlignFontProperty prop = new CT_VerticalAlignFontProperty();
            this.vertAlignField.Add(prop);
            return prop;
        }
        public void SetVertAlignArray(int index, CT_VerticalAlignFontProperty value)
        {
            this.vertAlignField[index] = value;
        }
        public void SetVertAlignArray(CT_VerticalAlignFontProperty[] array)
        {
            this.vertAlignField = new List<CT_VerticalAlignFontProperty>(array);
        }
        public CT_VerticalAlignFontProperty GetVertAlignArray(int index)
        {
            return this.vertAlignField[index];
        }
        #endregion vertAlign

        #region scheme
        [XmlElement]
        public List<CT_FontScheme> scheme
        {
            get { return this.schemeField; }
            set { this.schemeField = value; }
        }
        public int sizeOfSchemeArray()
        {
            return this.schemeField.Count;
        }
        public CT_FontScheme AddNewScheme()
        {
            CT_FontScheme newScheme = new CT_FontScheme();
            this.schemeField.Add(newScheme);
            return newScheme;
        }
        public void SetSchemeArray(int index, CT_FontScheme value)
        {
            this.schemeField[index] = value;
        }
        public CT_FontScheme GetSchemeArray(int index)
        {
            return this.schemeField[index];
        }
        #endregion scheme

        public override string ToString()
        {
            return CT_Font.GetString(this);
        }

        public CT_Font Clone()
        {
            CT_Font ctFont = new CT_Font();
            if (this.name.Count != 0)
            {
                CT_FontName newName = ctFont.AddNewName();
                newName.val = this.GetNameArray(0).val;
            }
            if (this.charset.Count != 0)
            {
                CT_IntProperty newCharset = ctFont.AddNewCharset();
                newCharset.val = this.GetCharsetArray(0).val;
            }
            if (this.family.Count != 0)
            {
                CT_IntProperty newFamily = ctFont.AddNewFamily();
                newFamily.val = this.GetFamilyArray(0).val;
            }
            if (this.b.Count != 0)
            {
                CT_BooleanProperty newB = ctFont.AddNewB();
                newB.val = this.GetBArray(0).val;
            }
            if (this.i.Count != 0)
            {
                CT_BooleanProperty newI = ctFont.AddNewI();
                newI.val = this.GetIArray(0).val;
            }
            if (this.strike.Count != 0)
            {
                CT_BooleanProperty newstrike = ctFont.AddNewStrike();
                newstrike.val = this.GetStrikeArray(0).val;
            }
            if (this.outline != null)
            {
                ctFont.outline = new CT_BooleanProperty();
                ctFont.outline.val = this.outline.val;
            }
            if (this.shadow != null)
            {
                ctFont.shadow = new CT_BooleanProperty();
                ctFont.shadow.val = this.shadow.val;
            }
            if (this.condense != null)
            {
                ctFont.condense = new CT_BooleanProperty();
                ctFont.condense.val = this.condense.val;
            }
            if (this.extend != null)
            {
                ctFont.extend = new CT_BooleanProperty();
                ctFont.extend.val = this.extend.val;
            }
            if (this.color.Count != 0)
            {
                CT_Color newColor = ctFont.AddNewColor();
                newColor.theme = this.GetColorArray(0).theme;
            }
            if (this.sz.Count != 0)
            {
                CT_FontSize newSz = ctFont.AddNewSz();
                newSz.val = this.GetSzArray(0).val;
            }
            if (this.u.Count != 0)
            {
                CT_UnderlineProperty newU = ctFont.AddNewU();
                newU.val = this.GetUArray(0).val;
            }
            if (this.vertAlign.Count != 0)
            {
                CT_VerticalAlignFontProperty newVertAlign = ctFont.AddNewVertAlign();
                newVertAlign.val = this.GetVertAlignArray(0).val;
            }
            if (this.scheme.Count != 0)
            {
                CT_FontScheme newFs = ctFont.AddNewScheme();
                newFs.val = this.GetSchemeArray(0).val;
            }
            return ctFont;
        }
    }
}
