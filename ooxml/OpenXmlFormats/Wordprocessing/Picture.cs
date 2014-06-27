using System;
using System.Xml;
using System.Xml.Serialization;
using System.Text;
using System.Collections.Generic;
using System.IO;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    [XmlInclude(typeof(CT_Picture))]
    [XmlInclude(typeof(CT_Object))]
    [XmlInclude(typeof(CT_Background))]

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PictureBase
    {

        private List<XmlNode> itemsField;

        private List<ItemsChoiceType9> itemsElementNameField;

        public CT_PictureBase()
        {
            this.itemsElementNameField = new List<ItemsChoiceType9>();
            this.itemsField = new List<XmlNode>();
        }

        [XmlAnyElement(Namespace = "urn:schemas-microsoft-com:office:office", Order = 0)]
        [XmlAnyElement(Namespace = "urn:schemas-microsoft-com:vml", Order = 0)]
        [XmlChoiceIdentifier("ItemsElementName")]
        public List<XmlNode> Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField =value;
            }
        }

        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
        public List<ItemsChoiceType9> ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
               this.itemsElementNameField = value;
            }
        }

        public void Set(object obj)
        {
            XmlSerializer xmlse = new XmlSerializer(obj.GetType());
            StringBuilder output = new StringBuilder();
            XmlWriterSettings settings = new XmlWriterSettings();

            settings.Encoding = Encoding.UTF8;
            settings.OmitXmlDeclaration = true;
            XmlWriter writer = XmlWriter.Create(output, settings);
            xmlse.Serialize(writer, obj);
            
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(output.ToString());
            lock(this)
            {
                this.itemsField.Add(xmlDoc.DocumentElement.CloneNode(true));
                this.itemsElementNameField.Add(ItemsChoiceType9.vml);
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Picture : CT_PictureBase
    {

        private CT_Rel movieField;

        private CT_Control controlField;

        public CT_Picture()
        {
            //this.controlField = new CT_Control();
            //this.movieField = new CT_Rel();
        }

        [XmlElement(Order = 0)]
        public CT_Rel movie
        {
            get
            {
                return this.movieField;
            }
            set
            {
                this.movieField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_Control control
        {
            get
            {
                return this.controlField;
            }
            set
            {
                this.controlField = value;
            }
        }

        // added because they are called in XWPFRun - perhaps another CT_Picture must be used instead
        public Dml.Picture.CT_PictureNonVisual AddNewNvPicPr()
        {
            throw new NotImplementedException();
        }

        public Dml.CT_BlipFillProperties AddNewBlipFill()
        {
            throw new NotImplementedException();
        }

        public Dml.CT_ShapeProperties AddNewSpPr()
        {
            throw new NotImplementedException();
        }

        public static CT_Picture Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Picture ctObj = new CT_Picture();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "movie")
                    ctObj.movie = CT_Rel.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "control")
                    ctObj.control = CT_Control.Parse(childNode, namespaceManager);
                else if(childNode.Prefix == "o")
                { 
                    ctObj.ItemsElementName.Add(ItemsChoiceType9.office);
                    ctObj.Items.Add(childNode);
                }else if(childNode.Prefix=="v")
                {
                    ctObj.ItemsElementName.Add(ItemsChoiceType9.vml);
                    ctObj.Items.Add(childNode);
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            sw.Write(">");
            if (this.movie != null)
                this.movie.Write(sw, "movie");
            if (this.control != null)
                this.control.Write(sw, "control");
            foreach (XmlNode childnode in Items)
            {
                sw.Write(childnode.OuterXml);
            }
            sw.Write(string.Format("</w:{0}>", nodeName));
        }
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Background : CT_PictureBase
    {

        private string colorField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;

        public CT_Background()
        {
            this.themeColorField = ST_ThemeColor.none;
        }

        public static CT_Background Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Background ctObj = new CT_Background();
            ctObj.color = XmlHelper.ReadString(node.Attributes["w:color"]);
            if (node.Attributes["w:themeColor"] != null)
                ctObj.themeColor = (ST_ThemeColor)Enum.Parse(typeof(ST_ThemeColor), node.Attributes["w:themeColor"].Value);
            ctObj.themeTint = XmlHelper.ReadBytes(node.Attributes["w:themeTint"]);
            ctObj.themeShade = XmlHelper.ReadBytes(node.Attributes["w:themeShade"]);

            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            if(themeColorField!= ST_ThemeColor.none)
                XmlHelper.WriteAttribute(sw, "w:themeColor", this.themeColor.ToString());
            XmlHelper.WriteAttribute(sw, "w:themeTint", this.themeTint);
            XmlHelper.WriteAttribute(sw, "w:themeShade", this.themeShade);
            XmlHelper.WriteAttribute(sw, "w:color", this.color);
            sw.Write("/>");
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string color
        {
            get
            {
                return this.colorField;
            }
            set
            {
                this.colorField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_ThemeColor themeColor
        {
            get
            {
                return this.themeColorField;
            }
            set
            {
                this.themeColorField = value;
            }
        }

        [XmlIgnore]
        public bool themeColorSpecified
        {
            get
            {
                return this.themeColorFieldSpecified;
            }
            set
            {
                this.themeColorFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeTint
        {
            get
            {
                return this.themeTintField;
            }
            set
            {
                this.themeTintField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] themeShade
        {
            get
            {
                return this.themeShadeField;
            }
            set
            {
                this.themeShadeField = value;
            }
        }
    }


    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Object : CT_PictureBase
    {

        private CT_Control controlField;

        private ulong dxaOrigField;

        private bool dxaOrigFieldSpecified;

        private ulong dyaOrigField;

        private bool dyaOrigFieldSpecified;

        public CT_Object()
        {
            //this.controlField = new CT_Control();
        }

        [XmlElement(Order = 0)]
        public CT_Control control
        {
            get
            {
                return this.controlField;
            }
            set
            {
                this.controlField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong dxaOrig
        {
            get
            {
                return this.dxaOrigField;
            }
            set
            {
                this.dxaOrigField = value;
            }
        }

        [XmlIgnore]
        public bool dxaOrigSpecified
        {
            get
            {
                return this.dxaOrigFieldSpecified;
            }
            set
            {
                this.dxaOrigFieldSpecified = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong dyaOrig
        {
            get
            {
                return this.dyaOrigField;
            }
            set
            {
                this.dyaOrigField = value;
            }
        }

        [XmlIgnore]
        public bool dyaOrigSpecified
        {
            get
            {
                return this.dyaOrigFieldSpecified;
            }
            set
            {
                this.dyaOrigFieldSpecified = value;
            }
        }
        public static CT_Object Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Object ctObj = new CT_Object();
            ctObj.dxaOrig = XmlHelper.ReadULong(node.Attributes["w:dxaOrig"]);
            ctObj.dyaOrig = XmlHelper.ReadULong(node.Attributes["w:dyaOrig"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "control")
                    ctObj.control = CT_Control.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:dxaOrig", this.dxaOrig);
            XmlHelper.WriteAttribute(sw, "w:dyaOrig", this.dyaOrig);
            sw.Write(">");
            if (this.control != null)
                this.control.Write(sw, "control");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Control
    {
        public static CT_Control Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Control ctObj = new CT_Control();
            ctObj.name = XmlHelper.ReadString(node.Attributes["w:name"]);
            ctObj.shapeid = XmlHelper.ReadString(node.Attributes["w:shapeid"]);
            ctObj.id = XmlHelper.ReadString(node.Attributes["r:id"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:name", this.name);
            XmlHelper.WriteAttribute(sw, "w:shapeid", this.shapeid);
            XmlHelper.WriteAttribute(sw, "r:id", this.id);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        private string nameField;

        private string shapeidField;

        private string idField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public string shapeid
        {
            get
            {
                return this.shapeidField;
            }
            set
            {
                this.shapeidField = value;
            }
        }

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }
    }

}
