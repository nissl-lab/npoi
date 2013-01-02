using System;
using System.Xml;
using System.Xml.Serialization;
using System.Text;
using System.Collections.Generic;

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

        private List<System.Xml.XmlElement> itemsField;

        private List<ItemsChoiceType9> itemsElementNameField;

        public CT_PictureBase()
        {
            this.itemsElementNameField = new List<ItemsChoiceType9>();
            this.itemsField = new List<System.Xml.XmlElement>();
        }

        [XmlAnyElement(Namespace = "urn:schemas-microsoft-com:office:office", Order = 0)]
        [XmlAnyElement(Namespace = "urn:schemas-microsoft-com:vml", Order = 0)]
        [XmlChoiceIdentifier("ItemsElementName")]
        public System.Xml.XmlElement[] Items
        {
            get
            {
                return this.itemsField.ToArray();
            }
            set
            {
                if (value == null)
                    this.itemsField = new List<XmlElement>();
                else
                    this.itemsField = new List<XmlElement>(value);
            }
        }

        [XmlElement("ItemsElementName", Order = 1)]
        [XmlIgnore]
        public ItemsChoiceType9[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField.ToArray();
            }
            set
            {
                if (value == null)
                    this.itemsElementNameField = new List<ItemsChoiceType9>();
                else
                    this.itemsElementNameField = new List<ItemsChoiceType9>(value);
            }
        }

        public void Set(object obj)
        {
            XmlSerializer xmlse = new XmlSerializer(obj.GetType());
            StringBuilder output = new StringBuilder();
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Encoding = Encoding.UTF8,
                OmitXmlDeclaration = true
            };
            XmlWriter writer = XmlWriter.Create(output, settings);
            xmlse.Serialize(writer, obj);
            
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(output.ToString());
            lock(this)
            {
                this.itemsField.Add((XmlElement)xmlDoc.DocumentElement.CloneNode(true));
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
            this.controlField = new CT_Control();
            this.movieField = new CT_Rel();
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
            this.controlField = new CT_Control();
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
    }

    [Serializable]

    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Control
    {

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
