
namespace NPOI.OpenXmlFormats
{
    using System.Xml.Serialization;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.ComponentModel;

    /// <remarks/>
    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")]
    [XmlRoot("Properties",Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", IsNullable = true)]
    public class CT_CustomProperties
    {
        public CT_CustomProperties()
        {
            propertyField = new List<CT_Property>();
        }
        private List<CT_Property> propertyField;

        /// <remarks/>
        [XmlElement("property")]
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
        public CT_CustomProperties Copy()
        {
            CT_CustomProperties prop = new CT_CustomProperties();
            prop.propertyField = new List<CT_Property>();
            foreach (CT_Property p in this.propertyField)
            {
                prop.propertyField.Add(p);
            }
            return prop;
        }
     }

    /// <remarks/>
    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", IsNullable = true)]
    public partial class CT_Property
    {

        private object itemField;

        private ItemChoiceType itemElementNameField;

        private string fmtidField;

        private int pidField;

        private string nameField;

        private string linkTargetField;

        /// <remarks/>
        [XmlElement("array", typeof(CT_Array), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("blob", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElement("bool", typeof(bool), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("bstr", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("cf", typeof(CT_Cf), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("clsid", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("cy", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("date", typeof(DateTime), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("decimal", typeof(decimal), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("empty", typeof(CT_Empty), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("error", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("filetime", typeof(System.DateTime), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("i1", typeof(sbyte), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("i2", typeof(short), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("i4", typeof(int), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("i8", typeof(long), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("int", typeof(int), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("lpstr", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("lpwstr", typeof(string), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("null", typeof(CT_Null), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("oblob", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElement("ostorage", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElement("ostream", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElement("r4", typeof(float), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("r8", typeof(double), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("storage", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElement("stream", typeof(byte[]), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", DataType = "base64Binary")]
        [XmlElement("ui1", typeof(byte), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("ui2", typeof(ushort), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("ui4", typeof(uint), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("ui8", typeof(ulong), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("uint", typeof(uint), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("vector", typeof(CT_Vector), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
        [XmlElement("vstream", typeof(CT_Vstream), Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
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
        [XmlIgnore]
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
        [XmlAttribute]
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
        [XmlAttribute]
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
        [XmlAttribute]
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
        [XmlAttribute]
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
        public override bool Equals(object obj)
        {
            if (!(obj is CT_Property))
                return false;

            CT_Property a = (CT_Property)obj;
            if (a.fmtidField != this.fmtidField
                ||a.itemElementNameField!=this.itemElementNameField
                ||a.itemField!=this.itemField
                ||a.linkTargetField!=this.linkTargetField
                ||a.nameField!=this.nameField
                ||a.pidField!=this.pidField)
                return false;

            return true;
        }
    }



    [System.Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
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

    [System.Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", IncludeInSchema = false)]
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