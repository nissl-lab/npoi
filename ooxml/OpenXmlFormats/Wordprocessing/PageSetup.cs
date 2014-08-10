using NPOI.OpenXml4Net.Util;
using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PageSz
    {

        private ulong wField;

        private bool wFieldSpecified;

        private ulong hField;

        private bool hFieldSpecified;

        private ST_PageOrientation orientField;

        private bool orientFieldSpecified;

        private string codeField;
        public static CT_PageSz Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PageSz ctObj = new CT_PageSz();
            ctObj.w = XmlHelper.ReadULong(node.Attributes["w:w"]);
            ctObj.h = XmlHelper.ReadULong(node.Attributes["w:h"]);
            if (node.Attributes["w:orient"] != null)
                ctObj.orient = (ST_PageOrientation)Enum.Parse(typeof(ST_PageOrientation), node.Attributes["w:orient"].Value);
            ctObj.code = XmlHelper.ReadString(node.Attributes["w:code"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:w", this.w);
            XmlHelper.WriteAttribute(sw, "w:h", this.h);
            if( this.orientField!= ST_PageOrientation.portrait)
                XmlHelper.WriteAttribute(sw, "w:orient", this.orient.ToString());
            XmlHelper.WriteAttribute(sw, "w:code", this.code);
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong w
        {
            get
            {
                return this.wField;
            }
            set
            {
                this.wField = value;
            }
        }

        [XmlIgnore]
        public bool wSpecified
        {
            get
            {
                return this.wFieldSpecified;
            }
            set
            {
                this.wFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong h
        {
            get
            {
                return this.hField;
            }
            set
            {
                this.hField = value;
            }
        }

        [XmlIgnore]
        public bool hSpecified
        {
            get
            {
                return this.hFieldSpecified;
            }
            set
            {
                this.hFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PageOrientation orient
        {
            get
            {
                return this.orientField;
            }
            set
            {
                this.orientField = value;
            }
        }

        [XmlIgnore]
        public bool orientSpecified
        {
            get
            {
                return this.orientFieldSpecified;
            }
            set
            {
                this.orientFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string code
        {
            get
            {
                return this.codeField;
            }
            set
            {
                this.codeField = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PageOrientation
    {

    
        portrait,

    
        landscape,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PageMar
    {

        private string topField;

        private ulong rightField;

        private string bottomField;

        private ulong leftField;

        private ulong headerField;

        private ulong footerField;

        private ulong gutterField;

        public static CT_PageMar Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PageMar ctObj = new CT_PageMar();
            ctObj.top = XmlHelper.ReadString(node.Attributes["w:top"]);
            ctObj.right = XmlHelper.ReadULong(node.Attributes["w:right"]);
            ctObj.bottom = XmlHelper.ReadString(node.Attributes["w:bottom"]);
            ctObj.left = XmlHelper.ReadULong(node.Attributes["w:left"]);
            ctObj.header = XmlHelper.ReadULong(node.Attributes["w:header"]);
            ctObj.footer = XmlHelper.ReadULong(node.Attributes["w:footer"]);
            ctObj.gutter = XmlHelper.ReadULong(node.Attributes["w:gutter"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:top", this.top);
            XmlHelper.WriteAttribute(sw, "w:right", this.right);
            XmlHelper.WriteAttribute(sw, "w:bottom", this.bottom);
            XmlHelper.WriteAttribute(sw, "w:left", this.left);
            XmlHelper.WriteAttribute(sw, "w:header", this.header);
            XmlHelper.WriteAttribute(sw, "w:footer", this.footer);
            XmlHelper.WriteAttribute(sw, "w:gutter", this.gutter);
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string top
        {
            get
            {
                return this.topField;
            }
            set
            {
                this.topField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string bottom
        {
            get
            {
                return this.bottomField;
            }
            set
            {
                this.bottomField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong header
        {
            get
            {
                return this.headerField;
            }
            set
            {
                this.headerField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong footer
        {
            get
            {
                return this.footerField;
            }
            set
            {
                this.footerField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong gutter
        {
            get
            {
                return this.gutterField;
            }
            set
            {
                this.gutterField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PaperSource
    {

        private string firstField;

        private string otherField;
        public static CT_PaperSource Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PaperSource ctObj = new CT_PaperSource();
            ctObj.first = XmlHelper.ReadString(node.Attributes["w:first"]);
            ctObj.other = XmlHelper.ReadString(node.Attributes["w:other"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:first", this.first);
            XmlHelper.WriteAttribute(sw, "w:other", this.other);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string first
        {
            get
            {
                return this.firstField;
            }
            set
            {
                this.firstField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string other
        {
            get
            {
                return this.otherField;
            }
            set
            {
                this.otherField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PageBorders
    {

        private CT_Border topField;

        private CT_Border leftField;

        private CT_Border bottomField;

        private CT_Border rightField;

        private ST_PageBorderZOrder zOrderField;

        private bool zOrderFieldSpecified;

        private ST_PageBorderDisplay displayField;

        private bool displayFieldSpecified;

        private ST_PageBorderOffset offsetFromField;

        private bool offsetFromFieldSpecified;

        public CT_PageBorders()
        {
            //this.rightField = new CT_Border();
            //this.bottomField = new CT_Border();
            //this.leftField = new CT_Border();
            //this.topField = new CT_Border();
        }
        public static CT_PageBorders Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PageBorders ctObj = new CT_PageBorders();
            if (node.Attributes["w:zOrder"] != null)
                ctObj.zOrder = (ST_PageBorderZOrder)Enum.Parse(typeof(ST_PageBorderZOrder), node.Attributes["w:zOrder"].Value);
            if (node.Attributes["w:display"] != null)
                ctObj.display = (ST_PageBorderDisplay)Enum.Parse(typeof(ST_PageBorderDisplay), node.Attributes["w:display"].Value);
            if (node.Attributes["w:offsetFrom"] != null)
                ctObj.offsetFrom = (ST_PageBorderOffset)Enum.Parse(typeof(ST_PageBorderOffset), node.Attributes["w:offsetFrom"].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "top")
                    ctObj.top = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "left")
                    ctObj.left = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "bottom")
                    ctObj.bottom = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "right")
                    ctObj.right = CT_Border.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:zOrder", this.zOrder.ToString());
            XmlHelper.WriteAttribute(sw, "w:display", this.display.ToString());
            XmlHelper.WriteAttribute(sw, "w:offsetFrom", this.offsetFrom.ToString());
            sw.Write(">");
            if (this.top != null)
                this.top.Write(sw, "top");
            if (this.left != null)
                this.left.Write(sw, "left");
            if (this.bottom != null)
                this.bottom.Write(sw, "bottom");
            if (this.right != null)
                this.right.Write(sw, "right");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        [XmlElement(Order = 0)]
        public CT_Border top
        {
            get
            {
                return this.topField;
            }
            set
            {
                this.topField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Border left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_Border bottom
        {
            get
            {
                return this.bottomField;
            }
            set
            {
                this.bottomField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_Border right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PageBorderZOrder zOrder
        {
            get
            {
                return this.zOrderField;
            }
            set
            {
                this.zOrderField = value;
            }
        }

        [XmlIgnore]
        public bool zOrderSpecified
        {
            get
            {
                return this.zOrderFieldSpecified;
            }
            set
            {
                this.zOrderFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PageBorderDisplay display
        {
            get
            {
                return this.displayField;
            }
            set
            {
                this.displayField = value;
            }
        }

        [XmlIgnore]
        public bool displaySpecified
        {
            get
            {
                return this.displayFieldSpecified;
            }
            set
            {
                this.displayFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_PageBorderOffset offsetFrom
        {
            get
            {
                return this.offsetFromField;
            }
            set
            {
                this.offsetFromField = value;
            }
        }

        [XmlIgnore]
        public bool offsetFromSpecified
        {
            get
            {
                return this.offsetFromFieldSpecified;
            }
            set
            {
                this.offsetFromFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PageBorderZOrder
    {

    
        front,

    
        back,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PageBorderDisplay
    {

    
        allPages,

    
        firstPage,

    
        notFirstPage,
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_PageBorderOffset
    {

    
        page,

    
        text,
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PageNumber
    {

        private ST_NumberFormat fmtField;

        private bool fmtFieldSpecified;

        private string startField;

        private string chapStyleField;

        private ST_ChapterSep chapSepField;

        private bool chapSepFieldSpecified;
        public static CT_PageNumber Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PageNumber ctObj = new CT_PageNumber();
            if (node.Attributes["w:fmt"] != null)
                ctObj.fmt = (ST_NumberFormat)Enum.Parse(typeof(ST_NumberFormat), node.Attributes["w:fmt"].Value);
            ctObj.start = XmlHelper.ReadString(node.Attributes["w:start"]);
            ctObj.chapStyle = XmlHelper.ReadString(node.Attributes["w:chapStyle"]);
            if (node.Attributes["w:chapSep"] != null)
                ctObj.chapSep = (ST_ChapterSep)Enum.Parse(typeof(ST_ChapterSep), node.Attributes["w:chapSep"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:fmt", this.fmt.ToString());
            XmlHelper.WriteAttribute(sw, "w:start", this.start);
            XmlHelper.WriteAttribute(sw, "w:chapStyle", this.chapStyle);
            XmlHelper.WriteAttribute(sw, "w:chapSep", this.chapSep.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_NumberFormat fmt
        {
            get
            {
                return this.fmtField;
            }
            set
            {
                this.fmtField = value;
            }
        }

        [XmlIgnore]
        public bool fmtSpecified
        {
            get
            {
                return this.fmtFieldSpecified;
            }
            set
            {
                this.fmtFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string start
        {
            get
            {
                return this.startField;
            }
            set
            {
                this.startField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string chapStyle
        {
            get
            {
                return this.chapStyleField;
            }
            set
            {
                this.chapStyleField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ChapterSep chapSep
        {
            get
            {
                return this.chapSepField;
            }
            set
            {
                this.chapSepField = value;
            }
        }

        [XmlIgnore]
        public bool chapSepSpecified
        {
            get
            {
                return this.chapSepFieldSpecified;
            }
            set
            {
                this.chapSepFieldSpecified = value;
            }
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SectType
    {

        private ST_SectionMark valField;

        private bool valFieldSpecified;
        public static CT_SectType Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SectType ctObj = new CT_SectType();
            if (node.Attributes["w:val"] != null)
                ctObj.val = (ST_SectionMark)Enum.Parse(typeof(ST_SectionMark), node.Attributes["w:val"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_SectionMark val
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

        [XmlIgnore]
        public bool valSpecified
        {
            get
            {
                return this.valFieldSpecified;
            }
            set
            {
                this.valFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_SectionMark
    {

    
        nextPage,

    
        nextColumn,

    
        continuous,

    
        evenPage,

    
        oddPage,
    }
    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LineNumber
    {

        private string countByField;

        private string startField;

        private ulong distanceField;

        private bool distanceFieldSpecified;

        private ST_LineNumberRestart restartField;

        private bool restartFieldSpecified;
        public static CT_LineNumber Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_LineNumber ctObj = new CT_LineNumber();
            ctObj.countBy = XmlHelper.ReadString(node.Attributes["w:countBy"]);
            ctObj.start = XmlHelper.ReadString(node.Attributes["w:start"]);
            ctObj.distance = XmlHelper.ReadULong(node.Attributes["w:distance"]);
            if (node.Attributes["w:restart"] != null)
                ctObj.restart = (ST_LineNumberRestart)Enum.Parse(typeof(ST_LineNumberRestart), node.Attributes["w:restart"].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:countBy", this.countBy);
            XmlHelper.WriteAttribute(sw, "w:start", this.start);
            XmlHelper.WriteAttribute(sw, "w:distance", this.distance);
            XmlHelper.WriteAttribute(sw, "w:restart", this.restart.ToString());
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string countBy
        {
            get
            {
                return this.countByField;
            }
            set
            {
                this.countByField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
        public string start
        {
            get
            {
                return this.startField;
            }
            set
            {
                this.startField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong distance
        {
            get
            {
                return this.distanceField;
            }
            set
            {
                this.distanceField = value;
            }
        }

        [XmlIgnore]
        public bool distanceSpecified
        {
            get
            {
                return this.distanceFieldSpecified;
            }
            set
            {
                this.distanceFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_LineNumberRestart restart
        {
            get
            {
                return this.restartField;
            }
            set
            {
                this.restartField = value;
            }
        }

        [XmlIgnore]
        public bool restartSpecified
        {
            get
            {
                return this.restartFieldSpecified;
            }
            set
            {
                this.restartFieldSpecified = value;
            }
        }
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_LineNumberRestart
    {

    
        newPage,

    
        newSection,

    
        continuous,
    }
}
