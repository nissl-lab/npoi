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
        //internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Font));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });

        // all elements are optional
        private CT_FontName nameField = null; // name of the font
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
        }

        public static CT_Font Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Font ctObj = new CT_Font();
            ctObj.charset = new List<CT_IntProperty>();
            ctObj.family = new List<CT_IntProperty>();
            ctObj.b = new List<CT_BooleanProperty>();
            ctObj.i = new List<CT_BooleanProperty>();
            ctObj.strike = new List<CT_BooleanProperty>();
            ctObj.color = new List<CT_Color>();
            ctObj.sz = new List<CT_FontSize>();
            ctObj.u = new List<CT_UnderlineProperty>();
            ctObj.vertAlign = new List<CT_VerticalAlignFontProperty>();
            ctObj.scheme = new List<CT_FontScheme>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "outline")
                    ctObj.outline = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shadow")
                    ctObj.shadow = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "condense")
                    ctObj.condense = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extend")
                    ctObj.extend = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "name")
                    ctObj.name= CT_FontName.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "charset")
                    ctObj.charset.Add(CT_IntProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "family")
                    ctObj.family.Add(CT_IntProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "b")
                    ctObj.b.Add(CT_BooleanProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "i")
                    ctObj.i.Add(CT_BooleanProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "strike")
                    ctObj.strike.Add(CT_BooleanProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "color")
                    ctObj.color.Add(CT_Color.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "sz")
                    ctObj.sz.Add(CT_FontSize.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "u")
                    ctObj.u.Add(CT_UnderlineProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "vertAlign")
                    ctObj.vertAlign.Add(CT_VerticalAlignFontProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "scheme")
                    ctObj.scheme.Add(CT_FontScheme.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.b != null)
            {
                foreach (CT_BooleanProperty x in this.b)
                {
                    x.Write(sw, "b");
                }
            }
            if (this.i != null)
            {
                foreach (CT_BooleanProperty x in this.i)
                {
                    x.Write(sw, "i");
                }
            }

            if (this.strike != null)
            {
                foreach (CT_BooleanProperty x in this.strike)
                {
                    x.Write(sw, "strike");
                }
            }
            if (this.condense != null)
                this.condense.Write(sw, "condense");
            if (this.extend != null)
                this.extend.Write(sw, "extend");
            if (this.outline != null)
                this.outline.Write(sw, "outline");
            if (this.shadow != null)
                this.shadow.Write(sw, "shadow");
            if (this.u != null)
            {
                foreach (CT_UnderlineProperty x in this.u)
                {
                    x.Write(sw, "u");
                }
            }
            if (this.vertAlign != null)
            {
                foreach (CT_VerticalAlignFontProperty x in this.vertAlign)
                {
                    x.Write(sw, "vertAlign");
                }
            }

            if (this.sz != null)
            {
                foreach (CT_FontSize x in this.sz)
                {
                    x.Write(sw, "sz");
                }
            }

            if (this.color != null)
            {
                foreach (CT_Color x in this.color)
                {
                    x.Write(sw, "color");
                }
            }
            if (this.name != null)
                this.name.Write(sw, "name");

            if (this.family != null)
            {
                foreach (CT_IntProperty x in this.family)
                {
                    x.Write(sw, "family");
                }
            }
            if (this.charset != null)
            {
                foreach (CT_IntProperty x in this.charset)
                {
                    x.Write(sw, "charset");
                }
            }
            if (this.scheme != null)
            {
                foreach (CT_FontScheme x in this.scheme)
                {
                    x.Write(sw, "scheme");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }


        //public static string GetString(CT_Font font)
        //{
        //    using (StringWriter writer = new StringWriter())
        //    {
        //        serializer.Serialize(writer, font, namespaces);
        //        return writer.ToString();
        //    }
        //}
        #region name
        [XmlElement]
        public CT_FontName name
        {
            get { return this.nameField; }
            set { this.nameField = value; }
        }
        public int sizeOfNameArray()
        {
            if (this.nameField == null)
                return 0;
            return 1;
        }
        public CT_FontName AddNewName()
        {
            this.nameField = new CT_FontName();
            return this.nameField;
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
            if (this.charsetField == null)
                return 0;
            return this.charsetField.Count;
        }
        public CT_IntProperty AddNewCharset()
        {
            if (this.charsetField == null)
                this.charsetField = new List<CT_IntProperty>();
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
            if (this.familyField == null)
                return 0;
            return this.familyField.Count;
        }
        public CT_IntProperty AddNewFamily()
        {
            if (this.familyField == null)
                this.familyField = new List<CT_IntProperty>();
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
        public int SizeOfBArray()
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
        public void SetBArray(List<CT_BooleanProperty> array)
        {
            this.bField = array;
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
            if (this.iField == null)
                return 0;
            return this.iField.Count;
        }
        public CT_BooleanProperty AddNewI()
        {
            if (this.iField == null)
                this.iField = new List<CT_BooleanProperty>();
            CT_BooleanProperty newI = new CT_BooleanProperty();
            this.iField.Add(newI);
            return newI;
        }
        public void SetIArray(int index, CT_BooleanProperty value)
        {
            this.iField[index] = value;
        }
        public void SetIArray(List<CT_BooleanProperty> array)
        {
            this.iField = array;
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
            if (this.strikeField == null)
                return 0;
            return this.strikeField.Count;
        }
        public CT_BooleanProperty AddNewStrike()
        {
            if (this.strikeField == null)
                this.strikeField = new List<CT_BooleanProperty>();
            CT_BooleanProperty prop = new CT_BooleanProperty();
            this.strikeField.Add(prop);
            return prop;
        }
        public void SetStrikeArray(int index, CT_BooleanProperty value)
        {
            this.strikeField[index] = value;
        }
        public void SetStrikeArray(List<CT_BooleanProperty> array)
        {
            this.strikeField = array;
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
        public int sizeOfShadowArray()
        {
            if (this.shadowField == null)
                return 0;
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
        public int sizeOfCondenseArray()
        {
            if (this.condenseField == null)
                return 0;
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
            if (this.colorField == null)
                return 0;
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
        public void SetColorArray(List<CT_Color> array)
        {
            this.colorField = array;
        }
        public CT_Color AddNewColor()
        {
            if (this.colorField == null)
                this.colorField = new List<CT_Color>();
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
            if (this.szField == null)
                return 0;
            return this.szField.Count;
        }
        public CT_FontSize AddNewSz()
        {
            if (this.szField == null)
                this.szField = new List<CT_FontSize>();
            CT_FontSize newFs = new CT_FontSize();
            this.szField.Add(newFs);
            return newFs;
        }
        public void SetSzArray(int index, CT_FontSize value)
        {
            this.szField[index] = value;
        }
        public void SetSzArray(List<CT_FontSize> array)
        {
            this.szField = array;
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
            if (this.uField == null)
                return 0;
            return this.uField.Count;
        }
        public CT_UnderlineProperty AddNewU()
        {
            if (this.uField == null)
                this.uField = new List<CT_UnderlineProperty>();
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
        public void SetUArray(List<CT_UnderlineProperty> array)
        {
            this.uField = array;
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
            if (this.vertAlignField == null)
                return 0;
            return this.vertAlignField.Count;
        }
        public CT_VerticalAlignFontProperty AddNewVertAlign()
        {
            if (this.vertAlignField == null)
                this.vertAlignField = new List<CT_VerticalAlignFontProperty>();
            CT_VerticalAlignFontProperty prop = new CT_VerticalAlignFontProperty();
            this.vertAlignField.Add(prop);
            return prop;
        }
        public void SetVertAlignArray(int index, CT_VerticalAlignFontProperty value)
        {
            this.vertAlignField[index] = value;
        }
        public void SetVertAlignArray(List<CT_VerticalAlignFontProperty> array)
        {
            this.vertAlignField =array;
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
            if (this.schemeField == null)
                return 0;
            return this.schemeField.Count;
        }
        public CT_FontScheme AddNewScheme()
        {
            if (this.schemeField == null)
                this.schemeField = new List<CT_FontScheme>();
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
            using (MemoryStream ms = new MemoryStream())
            {
                StreamWriter sw = new StreamWriter(ms);
                this.Write(sw, "font");
                sw.Flush();
                ms.Position = 0;
                StreamReader sr = new StreamReader(ms);
                string result = sr.ReadToEnd();
                return result;
            }
        }

        public CT_Font Clone()
        {
            CT_Font ctFont = new CT_Font();

            if (this.name!=null)
            {
                  CT_FontName newName = ctFont.AddNewName();
                    newName.val = this.name.val;
            }
            if (this.charset!=null)
            {
                foreach (CT_IntProperty ctCharset in this.charset)
                {
                    CT_IntProperty newCharset = ctFont.AddNewCharset();
                    newCharset.val = ctCharset.val;
                }
            }
            if (this.family!=null)
            {
                foreach (CT_IntProperty ctFamily in this.family)
                {
                    CT_IntProperty newFamily = ctFont.AddNewFamily();
                    newFamily.val = ctFamily.val;
                }
            }
            if (this.b != null)
            {
                foreach (CT_BooleanProperty ctB in this.b)
                {
                    CT_BooleanProperty newB = ctFont.AddNewB();
                    newB.val = ctB.val;
                }
            }
            if (this.i != null)
            {
                foreach (CT_BooleanProperty ctI in this.i)
                {
                    CT_BooleanProperty newI = ctFont.AddNewB();
                    newI.val = ctI.val;
                }
            }
            if (this.strike != null)
            {
                foreach (CT_BooleanProperty ctStrike in this.strike)
                {
                    CT_BooleanProperty newstrike = ctFont.AddNewStrike();
                    newstrike.val = ctStrike.val;
                }
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
            if (this.color != null)
            {
                foreach (CT_Color ctColor in this.color)
                {
                    CT_Color newColor = ctFont.AddNewColor();
                    newColor.theme = ctColor.theme;
                }
            }
            if (this.sz != null)
            {
                foreach (CT_FontSize ctSz in this.sz)
                {
                    CT_FontSize newSz = ctFont.AddNewSz();
                    newSz.val = ctSz.val;
                }
            }
            if (this.u != null)
            {
                foreach (CT_UnderlineProperty ctU in this.u)
                {
                    CT_UnderlineProperty newU = ctFont.AddNewU();
                    newU.val = ctU.val;
                }
            }
            if (this.vertAlign != null)
            {
                foreach (CT_VerticalAlignFontProperty ctVertAlign in this.vertAlign)
                {
                    CT_VerticalAlignFontProperty newVertAlign = ctFont.AddNewVertAlign();
                    newVertAlign.val = ctVertAlign.val;
                }

            }
            if (this.scheme != null)
            {
                foreach (CT_FontScheme ctScheme in this.scheme)
                {
                    CT_FontScheme newScheme = ctFont.AddNewScheme();
                    newScheme.val = ctScheme.val;
                }
            }
            return ctFont;
        }
    }
}
