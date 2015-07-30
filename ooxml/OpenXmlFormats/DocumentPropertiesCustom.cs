
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats
{
    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")]
    [XmlRoot("Properties", Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", IsNullable = true)]
    public class CT_CustomProperties
    {
        public CT_CustomProperties()
        {
            propertyField = new List<CT_Property>();
        }
        private List<CT_Property> propertyField;

    
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


    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", IsNullable = true)]
    public class CT_Property
    {

        private object itemField;

        private ItemChoiceType itemElementNameField;

        private string fmtidField;

        private int pidField;

        private string nameField;

        private string linkTargetField;

    
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
        [XmlChoiceIdentifier("ItemElementName")]
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
        public override int GetHashCode()
        {
            return this.pidField.GetHashCode();
        }
        public override string ToString()
        {
            return String.Format("[CT_Property][pid={0},name={1}]", pidField, nameField);
        }
        public bool IsSetLpwstr()
        {
            throw new NotImplementedException();
        }
    }



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")]
    public enum ST_ArrayBaseType
    {

    
        variant,

    
        i1,

    
        i2,

    
        i4,

    
        @int,

    
        ui1,

    
        ui2,

    
        ui4,

    
        @uint,

    
        r4,

    
        r8,

    
        @decimal,

    
        bstr,

    
        date,

    
        @bool,

    
        cy,

    
        error,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", IncludeInSchema = false)]
    public enum ItemChoiceType
    {

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:array")]
        array,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:blob")]
        blob,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:bool")]
        @bool,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:bstr")]
        bstr,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:cf")]
        cf,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:clsid")]
        clsid,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:cy")]
        cy,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:date")]
        date,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:decimal")]
        @decimal,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:empty")]
        empty,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:error")]
        error,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:filetime")]
        filetime,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:i1")]
        i1,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:i2")]
        i2,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:i4")]
        i4,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:i8")]
        i8,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:int")]
        @int,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:lpstr")]
        lpstr,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:lpwstr")]
        lpwstr,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:null")]
        @null,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:oblob")]
        oblob,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ostorage")]
        ostorage,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ostream")]
        ostream,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:r4")]
        r4,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:r8")]
        r8,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:storage")]
        storage,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:stream")]
        stream,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ui1")]
        ui1,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ui2")]
        ui2,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ui4")]
        ui4,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:ui8")]
        ui8,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:uint")]
        @uint,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:vector")]
        vector,

    
        [XmlEnum("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes:vstream")]
        vstream,
    }
}