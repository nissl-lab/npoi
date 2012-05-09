using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFTextType
    {

        private ST_FFTextType valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_FFTextType val
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_FFTextType
    {

    
        regular,

    
        number,

    
        date,

    
        currentTime,

    
        currentDate,

    
        calculated,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFName
    {

        private string valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FldChar
    {

        private object itemField;

        private ST_FldCharType fldCharTypeField;

        private ST_OnOff fldLockField;

        private bool fldLockFieldSpecified;

        private ST_OnOff dirtyField;

        private bool dirtyFieldSpecified;

        [System.Xml.Serialization.XmlElement("ffData", typeof(CT_FFData), Order = 0)]
        [System.Xml.Serialization.XmlElement("fldData", typeof(CT_Text), Order = 0)]
        [System.Xml.Serialization.XmlElement("numberingChange", typeof(CT_TrackChangeNumbering), Order = 0)]
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

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_FldCharType fldCharType
        {
            get
            {
                return this.fldCharTypeField;
            }
            set
            {
                this.fldCharTypeField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff fldLock
        {
            get
            {
                return this.fldLockField;
            }
            set
            {
                this.fldLockField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fldLockSpecified
        {
            get
            {
                return this.fldLockFieldSpecified;
            }
            set
            {
                this.fldLockFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_OnOff dirty
        {
            get
            {
                return this.dirtyField;
            }
            set
            {
                this.dirtyField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dirtySpecified
        {
            get
            {
                return this.dirtyFieldSpecified;
            }
            set
            {
                this.dirtyFieldSpecified = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFData
    {

        private object[] itemsField;

        private ItemsChoiceType14[] itemsElementNameField;

        public CT_FFData()
        {
            this.itemsElementNameField = new ItemsChoiceType14[0];
            this.itemsField = new object[0];
        }

        [System.Xml.Serialization.XmlElement("calcOnExit", typeof(CT_OnOff), Order = 0)]
        [System.Xml.Serialization.XmlElement("checkBox", typeof(CT_FFCheckBox), Order = 0)]
        [System.Xml.Serialization.XmlElement("ddList", typeof(CT_FFDDList), Order = 0)]
        [System.Xml.Serialization.XmlElement("enabled", typeof(CT_OnOff), Order = 0)]
        [System.Xml.Serialization.XmlElement("entryMacro", typeof(CT_MacroName), Order = 0)]
        [System.Xml.Serialization.XmlElement("exitMacro", typeof(CT_MacroName), Order = 0)]
        [System.Xml.Serialization.XmlElement("helpText", typeof(CT_FFHelpText), Order = 0)]
        [System.Xml.Serialization.XmlElement("name", typeof(CT_FFName), Order = 0)]
        [System.Xml.Serialization.XmlElement("statusText", typeof(CT_FFStatusText), Order = 0)]
        [System.Xml.Serialization.XmlElement("textInput", typeof(CT_FFTextInput), Order = 0)]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public object[] Items
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
        public ItemsChoiceType14[] ItemsElementName
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
    public class CT_FFCheckBox
    {

        private object itemField;

        private CT_OnOff defaultField;

        private CT_OnOff checkedField;

        public CT_FFCheckBox()
        {
            this.checkedField = new CT_OnOff();
            this.defaultField = new CT_OnOff();
        }

        [System.Xml.Serialization.XmlElement("size", typeof(CT_HpsMeasure), Order = 0)]
        [System.Xml.Serialization.XmlElement("sizeAuto", typeof(CT_OnOff), Order = 0)]
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

        [System.Xml.Serialization.XmlElement(Order = 1)]
        public CT_OnOff @default
        {
            get
            {
                return this.defaultField;
            }
            set
            {
                this.defaultField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 2)]
        public CT_OnOff @checked
        {
            get
            {
                return this.checkedField;
            }
            set
            {
                this.checkedField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFDDList
    {

        private CT_DecimalNumber resultField;

        private CT_DecimalNumber defaultField;

        private List<CT_String> listEntryField;

        public CT_FFDDList()
        {
            this.listEntryField = new List<CT_String>();
            this.defaultField = new CT_DecimalNumber();
            this.resultField = new CT_DecimalNumber();
        }

        [System.Xml.Serialization.XmlElement(Order = 0)]
        public CT_DecimalNumber result
        {
            get
            {
                return this.resultField;
            }
            set
            {
                this.resultField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 1)]
        public CT_DecimalNumber @default
        {
            get
            {
                return this.defaultField;
            }
            set
            {
                this.defaultField = value;
            }
        }

        [System.Xml.Serialization.XmlElement("listEntry", Order = 2)]
        public List<CT_String> listEntry
        {
            get
            {
                return this.listEntryField;
            }
            set
            {
                this.listEntryField = value;
            }
        }
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFHelpText
    {

        private ST_InfoTextType typeField;

        private bool typeFieldSpecified;

        private string valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_InfoTextType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool typeSpecified
        {
            get
            {
                return this.typeFieldSpecified;
            }
            set
            {
                this.typeFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_InfoTextType
    {

    
        text,

    
        autoText,
    }


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFStatusText
    {

        private ST_InfoTextType typeField;

        private bool typeFieldSpecified;

        private string valField;

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ST_InfoTextType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool typeSpecified
        {
            get
            {
                return this.typeFieldSpecified;
            }
            set
            {
                this.typeFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    [System.SerializableAttribute()]

    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [System.Xml.Serialization.XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_FFTextInput
    {

        private CT_FFTextType typeField;

        private CT_String defaultField;

        private CT_DecimalNumber maxLengthField;

        private CT_String formatField;

        public CT_FFTextInput()
        {
            this.formatField = new CT_String();
            this.maxLengthField = new CT_DecimalNumber();
            this.defaultField = new CT_String();
            this.typeField = new CT_FFTextType();
        }

        [System.Xml.Serialization.XmlElement(Order = 0)]
        public CT_FFTextType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 1)]
        public CT_String @default
        {
            get
            {
                return this.defaultField;
            }
            set
            {
                this.defaultField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 2)]
        public CT_DecimalNumber maxLength
        {
            get
            {
                return this.maxLengthField;
            }
            set
            {
                this.maxLengthField = value;
            }
        }

        [System.Xml.Serialization.XmlElement(Order = 3)]
        public CT_String format
        {
            get
            {
                return this.formatField;
            }
            set
            {
                this.formatField = value;
            }
        }
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType14
    {

    
        calcOnExit,

    
        checkBox,

    
        ddList,

    
        enabled,

    
        entryMacro,

    
        exitMacro,

    
        helpText,

    
        name,

    
        statusText,

    
        textInput,
    }


    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_FldCharType
    {

    
        begin,

    
        separate,

    
        end,
    }

}
