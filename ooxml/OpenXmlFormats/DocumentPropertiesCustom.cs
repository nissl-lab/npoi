
namespace NPOI.OpenXmlFormats
{
    using System.Xml.Serialization;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.ComponentModel;

    /// <remarks/>
    [SerializableAttribute()]
    [DebuggerStepThroughAttribute()]
    [DesignerCategoryAttribute("code")]
    [XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")]
    [XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", IsNullable = true)]
    public partial class CT_Properties
    {
        public CT_Properties()
        {
            propertyField = new List<CT_Property>();
        }
        private List<CT_Property> propertyField;

        /// <remarks/>
        [XmlElementAttribute("property")]
        public List<CT_Property> property
        {
            get
            {
                return this.propertyField;
            }
            set
            {
                this.propertyField = value;
            }
        }
        public int sizeOfPropertyArray()
        {
            return this.propertyField.Count;
        }
        public CT_Property AddNewProperty()
        {
            CT_Property p = new CT_Property();
            propertyField.Add(p);
            return p;
        }
        public CT_Property GetPropertyArray(int index)
        {
            return this.propertyField[index];
        }
        public List<CT_Property> GetPropertyList()
        {
            return propertyField;
        }
        public CT_Property GetProperty(string name)
        {
            for (int i = 0; i < propertyField.Count; i++)
            {
                if (propertyField[i].name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                {
                    return propertyField[i];
                }
            }
            return null;
        }
     }

    /// <remarks/>
    [SerializableAttribute()]
    [DebuggerStepThroughAttribute()]
    [DesignerCategoryAttribute("code")]
    [XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")]
    [XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", IsNullable = true)]
    public partial class CT_Property
    {

        private object itemField;

        private ItemChoiceType itemElementNameField;

        private string fmtidField;

        private int pidField;

        private string nameField;

        private string linkTargetField;

        /// <remarks/>
        [XmlElementAttribute("array", typeof(CT_Array), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("blob", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElementAttribute("bool", typeof(bool), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("bstr", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("cf", typeof(CT_Cf), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("clsid", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("cy", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("date", typeof(DateTime), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("decimal", typeof(decimal), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("empty", typeof(CT_Empty), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("error", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("filetime", typeof(System.DateTime), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("i1", typeof(sbyte), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("i2", typeof(short), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("i4", typeof(int), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("i8", typeof(long), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("int", typeof(int), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("lpstr", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("lpwstr", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("null", typeof(CT_Null), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("oblob", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElementAttribute("ostorage", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElementAttribute("ostream", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElementAttribute("r4", typeof(float), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("r8", typeof(double), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("storage", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElementAttribute("stream", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElementAttribute("ui1", typeof(byte), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("ui2", typeof(ushort), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("ui4", typeof(uint), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("ui8", typeof(ulong), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("uint", typeof(uint), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("vector", typeof(CT_Vector), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElementAttribute("vstream", typeof(CT_Vstream), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlChoiceIdentifierAttribute("ItemElementName")]
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

        /// <remarks/>
        [XmlIgnoreAttribute()]
        public ItemChoiceType ItemElementName
        {
            get
            {
                return this.itemElementNameField;
            }
            set
            {
                this.itemElementNameField = value;
            }
        }

        /// <remarks/>
        [XmlAttributeAttribute()]
        public string fmtid
        {
            get
            {
                return this.fmtidField;
            }
            set
            {
                this.fmtidField = value;
            }
        }

        /// <remarks/>
        [XmlAttributeAttribute()]
        public int pid
        {
            get
            {
                return this.pidField;
            }
            set
            {
                this.pidField = value;
            }
        }

        /// <remarks/>
        [XmlAttributeAttribute()]
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

        /// <remarks/>
        [XmlAttributeAttribute()]
        public string linkTarget
        {
            get
            {
                return this.linkTargetField;
            }
            set
            {
                this.linkTargetField = value;
            }
        }
    }



    [System.SerializableAttribute()]
    [XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
    public enum ST_ArrayBaseType
    {

        /// <remarks/>
        variant,

        /// <remarks/>
        i1,

        /// <remarks/>
        i2,

        /// <remarks/>
        i4,

        /// <remarks/>
        @int,

        /// <remarks/>
        ui1,

        /// <remarks/>
        ui2,

        /// <remarks/>
        ui4,

        /// <remarks/>
        @uint,

        /// <remarks/>
        r4,

        /// <remarks/>
        r8,

        /// <remarks/>
        @decimal,

        /// <remarks/>
        bstr,

        /// <remarks/>
        date,

        /// <remarks/>
        @bool,

        /// <remarks/>
        cy,

        /// <remarks/>
        error,
    }

    [System.SerializableAttribute()]
    [XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", IncludeInSchema = false)]
    public enum ItemChoiceType
    {

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:array")]
        array,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:blob")]
        blob,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:bool")]
        @bool,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:bstr")]
        bstr,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:cf")]
        cf,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:clsid")]
        clsid,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:cy")]
        cy,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:date")]
        date,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:decimal")]
        @decimal,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:empty")]
        empty,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:error")]
        error,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:filetime")]
        filetime,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:i1")]
        i1,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:i2")]
        i2,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:i4")]
        i4,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:i8")]
        i8,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:int")]
        @int,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:lpstr")]
        lpstr,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:lpwstr")]
        lpwstr,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:null")]
        @null,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:oblob")]
        oblob,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ostorage")]
        ostorage,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ostream")]
        ostream,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:r4")]
        r4,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:r8")]
        r8,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:storage")]
        storage,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:stream")]
        stream,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ui1")]
        ui1,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ui2")]
        ui2,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ui4")]
        ui4,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ui8")]
        ui8,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:uint")]
        @uint,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:vector")]
        vector,

        /// <remarks/>
        [XmlEnumAttribute("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:vstream")]
        vstream,
    }
}