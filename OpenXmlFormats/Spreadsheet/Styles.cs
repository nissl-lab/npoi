
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Spreadsheet
{

    [Serializable]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "styleSheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class CT_Stylesheet
    {

        private CT_NumFmts numFmtsField;

        private CT_Fonts fontsField;

        private CT_Fills fillsField;

        private CT_Borders bordersField;

        private CT_CellStyleXfs cellStyleXfsField;

        private CT_CellXfs cellXfsField;

        private CT_CellStyles cellStylesField;

        private CT_Dxfs dxfsField;

        private CT_TableStyles tableStylesField;

        private CT_Colors colorsField;

        private CT_ExtensionList extLstField;

        public CT_Stylesheet()
        {
            //this.extLstField = new CT_ExtensionList();
            //this.colorsField = new CT_Colors();
            //this.tableStylesField = new CT_TableStyles();
            //this.dxfsField = new CT_Dxfs();
            //this.cellStylesField = new CT_CellStyles();
            //this.bordersField = new CT_Borders();
            //this.fillsField = new CT_Fills();
            //this.fontsField = new CT_Fonts();
            //this.numFmtsField = new CT_NumFmts();
        }
        public static CT_Stylesheet Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Stylesheet ctObj = new CT_Stylesheet();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "numFmts")
                    ctObj.numFmts = CT_NumFmts.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fonts")
                    ctObj.fonts = CT_Fonts.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fills")
                    ctObj.fills = CT_Fills.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "borders")
                    ctObj.borders = CT_Borders.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cellStyleXfs")
                    ctObj.cellStyleXfs = CT_CellStyleXfs.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cellXfs")
                    ctObj.cellXfs = CT_CellXfs.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cellStyles")
                    ctObj.cellStyles = CT_CellStyles.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "dxfs")
                    ctObj.dxfs = CT_Dxfs.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "tableStyles")
                    ctObj.tableStyles = CT_TableStyles.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "colors")
                    ctObj.colors = CT_Colors.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            sw.Write(" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac x16r2 xr\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\"");
            sw.Write(">");
            if (this.numFmts != null)
                this.numFmts.Write(sw, "numFmts");
            if (this.fonts != null)
                this.fonts.Write(sw, "fonts");
            if (this.fills != null)
                this.fills.Write(sw, "fills");
            if (this.borders != null)
                this.borders.Write(sw, "borders");
            if (this.cellStyleXfs != null)
                this.cellStyleXfs.Write(sw, "cellStyleXfs");
            if (this.cellXfs != null)
                this.cellXfs.Write(sw, "cellXfs");
            if (this.cellStyles != null)
                this.cellStyles.Write(sw, "cellStyles");
            if (this.dxfs != null)
                this.dxfs.Write(sw, "dxfs");
            if (this.tableStyles != null)
                this.tableStyles.Write(sw, "tableStyles");
            if (this.colors != null)
                this.colors.Write(sw, "colors");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            sw.Write("</styleSheet>");
        }

        public CT_Borders AddNewBorders()
        {
            this.bordersField = new CT_Borders();
            return this.bordersField;
        }
        public CT_CellStyleXfs AddNewCellStyleXfs()
        {
            this.cellStyleXfsField = new CT_CellStyleXfs();
            return this.cellStyleXfsField;
        }
        public CT_CellXfs AddNewCellXfs()
        {
            this.cellXfsField = new CT_CellXfs();
            return this.cellXfsField;
        }
        [XmlElement]
        public CT_NumFmts numFmts
        {
            get
            {
                return this.numFmtsField;
            }
            set
            {
                this.numFmtsField = value;
            }
        }
        [XmlElement]
        public CT_Fonts fonts
        {
            get
            {
                return this.fontsField;
            }
            set
            {
                this.fontsField = value;
            }
        }
        [XmlElement]
        public CT_Fills fills
        {
            get
            {
                return this.fillsField;
            }
            set
            {
                this.fillsField = value;
            }
        }
        [XmlElement]
        public CT_Borders borders
        {
            get
            {
                return this.bordersField;
            }
            set
            {
                this.bordersField = value;
            }
        }
        [XmlElement]
        public CT_CellStyleXfs cellStyleXfs
        {
            get
            {
                return this.cellStyleXfsField;
            }
            set
            {
                this.cellStyleXfsField = value;
            }
        }
        [XmlElement]
        public CT_CellXfs cellXfs
        {
            get
            {
                return this.cellXfsField;
            }
            set
            {
                this.cellXfsField = value;
            }
        }
        [XmlElement]
        public CT_CellStyles cellStyles
        {
            get
            {
                return this.cellStylesField;
            }
            set
            {
                this.cellStylesField = value;
            }
        }
        [XmlElement]
        public CT_Dxfs dxfs
        {
            get
            {
                return this.dxfsField;
            }
            set
            {
                this.dxfsField = value;
            }
        }
        [XmlElement]
        public CT_TableStyles tableStyles
        {
            get
            {
                return this.tableStylesField;
            }
            set
            {
                this.tableStylesField = value;
            }
        }
        [XmlElement]
        public CT_Colors colors
        {
            get
            {
                return this.colorsField;
            }
            set
            {
                this.colorsField = value;
            }
        }
        [XmlElement]
        public CT_ExtensionList extLst
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










    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_FontScheme
    {
        private ST_FontScheme valField;

    
        [XmlAttribute]
        public ST_FontScheme val
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

        public static CT_FontScheme Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FontScheme ctObj = new CT_FontScheme();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_FontScheme)Enum.Parse(typeof(ST_FontScheme), node.Attributes["val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val.ToString());
            sw.Write("/>");
        }


    }



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_FontName
    {

        private string valField;

        public static CT_FontName Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FontName ctObj = new CT_FontName();
            ctObj.val = XmlHelper.ReadString(node.Attributes["val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val);
            sw.Write("/>");
        }


        [XmlAttribute]
        public string val
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_FontSize
    {
        private double valField;
        public static CT_FontSize Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FontSize ctObj = new CT_FontSize();
            ctObj.val = XmlHelper.ReadDouble(node.Attributes["val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val);
            sw.Write("/>");
        }



        [XmlAttribute]
        public double val
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_FontScheme
    {
    
        none = 1,

    
        major = 2,

    
        minor = 3,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_UnderlineProperty
    {

        private ST_UnderlineValues? valField = null;

        public CT_UnderlineProperty()
        {
            this.valField = ST_UnderlineValues.single;
        }

        public static CT_UnderlineProperty Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_UnderlineProperty ctObj = new CT_UnderlineProperty();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_UnderlineValues)Enum.Parse(typeof(ST_UnderlineValues), node.Attributes["val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val.ToString());
            sw.Write("/>");
        }




        [DefaultValue(ST_UnderlineValues.single)]
        [XmlAttribute]
        public ST_UnderlineValues val
        {
            get
            {
                return  (null == valField) ? ST_UnderlineValues.single : (ST_UnderlineValues)this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
        [XmlIgnore]
        public bool valbSpecified
        {
            get { return (null != valField); }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_UnderlineValues
    {
    
        none,

    
        single,

    
        [XmlEnum("double")]
        @double,

    
        singleAccounting,

    
        doubleAccounting,

    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_VerticalAlignFontProperty
    {
        private ST_VerticalAlignRun valField;

        [XmlAttribute]
        public ST_VerticalAlignRun val
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
        public static CT_VerticalAlignFontProperty Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_VerticalAlignFontProperty ctObj = new CT_VerticalAlignFontProperty();
            if (node.Attributes["val"] != null)
                ctObj.val = (ST_VerticalAlignRun)Enum.Parse(typeof(ST_VerticalAlignRun), node.Attributes["val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val.ToString());
            sw.Write("/>");
        }


    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_VerticalAlignRun
    {
    
        baseline,

    
        superscript,

    
        subscript,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_BooleanProperty
    {

        private bool valField = true;

        public CT_BooleanProperty()
        {
        }
        public static CT_BooleanProperty Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_BooleanProperty ctObj = new CT_BooleanProperty();
            if (node.Attributes["val"]!=null)
                ctObj.val = XmlHelper.ReadBool(node.Attributes["val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            if(!val)
                XmlHelper.WriteAttribute(sw, "val", this.val);
            sw.Write("/>");
        }


        [DefaultValue(true)]
        [XmlAttribute]
        public bool val
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
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_IntProperty
    {

        private int valField;

        [XmlAttribute]
        public int val
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

        public static CT_IntProperty Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_IntProperty ctObj = new CT_IntProperty();
            ctObj.val = XmlHelper.ReadInt(node.Attributes["val"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "val", this.val);
            sw.Write("/>");
        }
    }


}
