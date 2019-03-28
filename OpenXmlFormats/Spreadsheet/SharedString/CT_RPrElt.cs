using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Properties of Rich Text Run.
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_RPrElt
    {
        // all elements are optional
        private CT_FontName rFontField = null; // name of the font
        private CT_IntProperty charsetField = null;
        private CT_IntProperty familyField = null; // family of the font
        private CT_BooleanProperty bField = null; // typeface bold
        private CT_BooleanProperty iField = null;   // italic
        private CT_BooleanProperty strikeField = null; //   strike through
        private CT_BooleanProperty outlineField = null;
        private CT_BooleanProperty shadowField = null;
        private CT_BooleanProperty condenseField = null;
        private CT_BooleanProperty extendField = null;
        private CT_Color colorField = null;
        private CT_FontSize szField = null; // size of the font
        private CT_UnderlineProperty uField = null; // underline
        private CT_VerticalAlignFontProperty vertAlignField = null;  // vertical alignment of the text
        private CT_FontScheme schemeField = null;

        public static CT_RPrElt Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RPrElt ctObj = new CT_RPrElt();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "rFont")
                    ctObj.rFont = CT_FontName.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "charset")
                    ctObj.charset = CT_IntProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "family")
                    ctObj.family = CT_IntProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "b")
                    ctObj.b = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "i")
                    ctObj.i = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "strike")
                    ctObj.strike = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "outline")
                    ctObj.outline = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shadow")
                    ctObj.shadow = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "condense")
                    ctObj.condense = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extend")
                    ctObj.extend = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "color")
                    ctObj.color = CT_Color.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "sz")
                    ctObj.sz = CT_FontSize.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "u")
                    ctObj.u = CT_UnderlineProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "vertAlign")
                    ctObj.vertAlign = CT_VerticalAlignFontProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "scheme")
                    ctObj.scheme = CT_FontScheme.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.sz != null)
                this.sz.Write(sw, "sz");
            if (this.color != null)
                this.color.Write(sw, "color");
            if (this.rFont != null)
                this.rFont.Write(sw, "rFont");
            if (this.family != null)
                this.family.Write(sw, "family");
            if (this.charset != null)
                this.charset.Write(sw, "charset");
            if (this.b != null)
                this.b.Write(sw, "b");
            if (this.i != null)
                this.i.Write(sw, "i");
            if (this.strike != null)
                this.strike.Write(sw, "strike");
            if (this.outline != null)
                this.outline.Write(sw, "outline");
            if (this.shadow != null)
                this.shadow.Write(sw, "shadow");
            if (this.condense != null)
                this.condense.Write(sw, "condense");
            if (this.extend != null)
                this.extend.Write(sw, "extend");
            if (this.u != null)
                this.u.Write(sw, "u");
            if (this.vertAlign != null)
                this.vertAlign.Write(sw, "vertAlign");
            if (this.scheme != null)
                this.scheme.Write(sw, "scheme");
            sw.Write(string.Format("</{0}>", nodeName));
        }


        #region rFont
        [XmlElement]
        public CT_FontName rFont
        {
            get { return this.rFontField; }
            set { this.rFontField = value; }
        }
        [XmlIgnore]
        // do not remove this field or change the name, because it is automatically used by the XmlSerializer to decide if the name attribute should be printed or not.
        public bool rFontSpecified
        {
            get { return (null != rFontField); }
        }
        public int SizeOfRFontArray()
        {
            return this.rFontField == null ? 0 : 1;
        }
        public CT_FontName AddNewRFont()
        {
            this.rFontField = new CT_FontName();
            return this.rFontField;
        }
        public CT_FontName GetRFontArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.rFontField;
        }
        #endregion rFont

        #region charset
        [XmlElement]
        public CT_IntProperty charset
        {
            get { return this.charsetField; }
            set { this.charsetField = value; }
        }
        [XmlIgnore]
        public bool charsetSpecified
        {
            get { return (null != charsetField); }
        }
        public int sizeOfCharsetArray()
        {
            return this.charsetField == null ? 0 : 1;
        }
        public CT_IntProperty AddNewCharset()
        {
            this.charsetField = new CT_IntProperty();
            return this.charsetField;
        }
        public CT_IntProperty GetCharsetArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.charsetField;
        }
        #endregion charset

        #region family
        [XmlElement]
        public CT_IntProperty family
        {
            get { return this.familyField; }
            set { this.familyField = value; }
        }
        [XmlIgnore]
        public bool familySpecified
        {
            get { return (null != familyField); }
        }
        public int SizeOfFamilyArray()
        {
            return this.familyField == null ? 0 : 1;
        }
        public CT_IntProperty AddNewFamily()
        {
            this.familyField = new CT_IntProperty();
            return this.familyField;
        }
        //public void SetFamilyArray()
        //{
        //    this.familyField = null;
        //}
        public CT_IntProperty GetFamilyArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.familyField;
        }
        #endregion family

        #region b
        [XmlElement]
        public CT_BooleanProperty b
        {
            get { return this.bField; }
            set { this.bField = value; }
        }
        [XmlIgnore]
        public bool bSpecified
        {
            get { return (null != bField); }
        }
        public int SizeOfBArray()
        {
            return this.bField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewB()
        {
            this.bField = new CT_BooleanProperty();
            return this.bField;
        }
        public void SetBArray(CT_BooleanProperty[] array)
        {
            this.bField = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetBArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.bField;
        }
        #endregion b

        #region i
        [XmlElement]
        public CT_BooleanProperty i
        {
            get { return this.iField; }
            set { this.iField = value; }
        }
        [XmlIgnore]
        public bool iSpecified
        {
            get { return (null != iField); }
        }
        public int SizeOfIArray()
        {
            return this.iField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewI()
        {
            this.iField = new CT_BooleanProperty();
            return this.iField;
        }
        public void SetIArray(CT_BooleanProperty[] array)
        {
            this.iField = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetIArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.iField;
        }
        #endregion i

        #region strike
        [XmlElement]
        public CT_BooleanProperty strike
        {
            get { return this.strikeField; }
            set { this.strikeField = value; }
        }
        [XmlIgnore]
        public bool strikeSpecified
        {
            get { return (null != strikeField); }
        }
        public int sizeOfStrikeArray()
        {
            return this.strikeField == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewStrike()
        {
            this.strikeField = new CT_BooleanProperty();
            return this.strikeField;
        }
        public void SetStrikeArray(CT_BooleanProperty[] array)
        {
            this.strikeField = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetStrikeArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.strikeField;
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
        public CT_Color color
        {
            get { return this.colorField; }
            set { this.colorField = value; }
        }
        [XmlIgnore]
        public bool colorSpecified
        {
            get { return (null != colorField); }
        }
        public int SizeOfColorArray()
        {
            return this.colorField == null ? 0 : 1;
        }
        public CT_Color GetColorArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.colorField;
        }
        public void SetColorArray(CT_Color[] array)
        {
            this.colorField = array.Length > 0 ? array[0] : null;
        }
        public CT_Color AddNewColor()
        {
            this.colorField = new CT_Color();
            return this.colorField;
        }
        #endregion color

        #region sz
        [XmlElement]
        public CT_FontSize sz
        {
            get { return this.szField; }
            set { this.szField = value; }
        }
        [XmlIgnore]
        public bool szSpecified
        {
            get { return (null != szField); }
        }
        public int SizeOfSzArray()
        {
            return this.szField == null ? 0 : 1;
        }
        public CT_FontSize AddNewSz()
        {
            this.szField = new CT_FontSize();
            return this.szField;
        }
        public void SetSzArray(CT_FontSize[] array)
        {
            this.szField = array.Length > 0 ? array[0] : null;
        }
        public CT_FontSize GetSzArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.szField;
        }
        #endregion sz

        #region u
        [XmlElement]
        public CT_UnderlineProperty u
        {
            get { return this.uField; }
            set { this.uField = value; }
        }
        [XmlIgnore]
        public bool uSpecified
        {
            get { return (null != uField); }
        }
        public int SizeOfUArray()
        {
            return this.uField == null ? 0 : 1;
        }
        public CT_UnderlineProperty AddNewU()
        {
            this.uField = new CT_UnderlineProperty();
            return this.uField;
        }
        public void SetUArray(CT_UnderlineProperty[] array)
        {
            this.uField = array.Length > 0 ? array[0] : null;
        }
        public CT_UnderlineProperty GetUArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.uField;
        }
        #endregion u

        #region vertAlign
        [XmlElement]
        public CT_VerticalAlignFontProperty vertAlign
        {
            get { return this.vertAlignField; }
            set { this.vertAlignField = value; }
        }
        [XmlIgnore]
        public bool vertAlignSpecified
        {
            get { return (null != vertAlignField); }
        }
        public int sizeOfVertAlignArray()
        {
            return this.vertAlignField == null ? 0 : 1;
        }
        public CT_VerticalAlignFontProperty AddNewVertAlign()
        {
            this.vertAlignField = new CT_VerticalAlignFontProperty();
            return this.vertAlignField;
        }
        public void SetVertAlignArray(CT_VerticalAlignFontProperty[] array)
        {
            this.vertAlignField = array.Length > 0 ? array[0] : null;
        }
        public CT_VerticalAlignFontProperty GetVertAlignArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.vertAlignField;
        }
        #endregion vertAlign

        #region scheme
        [XmlElement]
        public CT_FontScheme scheme
        {
            get { return this.schemeField; }
            set { this.schemeField = value; }
        }
        [XmlIgnore]
        public bool schemeSpecified
        {
            get { return (null != schemeField); }
        }
        public int sizeOfSchemeArray()
        {
            return this.schemeField == null ? 0 : 1;
        }
        public CT_FontScheme AddNewScheme()
        {
            this.schemeField = new CT_FontScheme();
            return this.schemeField;
        }
        public CT_FontScheme GetSchemeArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.schemeField;
        }
        #endregion scheme

        //public int SizeOfUArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfBArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfSzArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfRFontArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfIArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfColorArray()
        //{
        //    throw new NotImplementedException();
        //}
    }
}
