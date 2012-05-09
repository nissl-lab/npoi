using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NPOI.OpenXmlFormats.Shared;

namespace NPOI.OpenXmlFormats.Wordprocessing
{
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(CT_Picture))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(CT_Object))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(CT_Background))]

    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PictureBase
    {

        private System.Xml.XmlElement[] itemsField;

        private ItemsChoiceType9[] itemsElementNameField;

        public CT_PictureBase()
        {
            this.itemsElementNameField = new ItemsChoiceType9[0];
            this.itemsField = new System.Xml.XmlElement[0];
        }

        [System.Xml.Serialization.XmlAnyElementAttribute(Namespace = "urn:schemas-microsoft-com:office:office", Order = 0)]
        [System.Xml.Serialization.XmlAnyElementAttribute(Namespace = "urn:schemas-microsoft-com:vml", Order = 0)]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public System.Xml.XmlElement[] Items
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

        [System.Xml.Serialization.XmlElement("ItemsElementName", Order = 1)]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public ItemsChoiceType9[] ItemsElementName
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
    }

    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Picture : CT_PictureBase
    {

        private CT_Rel movieField;

        private CT_Control controlField;

        public CT_Picture()
        {
            this.controlField = new CT_Control();
            this.movieField = new CT_Rel();
        }

        [System.Xml.Serialization.XmlElement(Order = 0)]
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

        [System.Xml.Serialization.XmlElement(Order = 1)]
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

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Background : CT_PictureBase
    {

        private string colorField;

        private ST_ThemeColor themeColorField;

        private bool themeColorFieldSpecified;

        private byte[] themeTintField;

        private byte[] themeShadeField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
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

        [System.Xml.Serialization.XmlElement(Order = 0)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Control
    {

        private string nameField;

        private string shapeidField;

        private string idField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
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
