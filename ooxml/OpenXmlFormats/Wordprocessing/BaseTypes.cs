using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Wordprocessing
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Empty
    {
    }

    /// <summary>
    /// On/Off Value
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_OnOff
    {
        public static CT_OnOff Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_OnOff ctObj = new CT_OnOff();
            if (node.Attributes["w:val"] != null)
            {
                ctObj.valField = XmlHelper.ReadBool(node.Attributes["w:val"]);
            }
            else
            {
                ctObj.valField = true;
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            //if true, don't render val attribute
            if (!this.valField)
                XmlHelper.WriteAttribute(sw, "w:val", this.valField, true);
            sw.Write("/>");
        }


        private bool valField;

        private bool valFieldSpecified;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public bool val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
                this.valFieldSpecified = true;
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

        public void UnSetVal()
        {
            this.val = false;
            this.valFieldSpecified = false;
        }

        public bool IsSetVal()
        {
            return this.valFieldSpecified;
        }
    }
    /// <summary>
    /// On/Off Value
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    public enum ST_OnOff
    {
        /// <summary>
        /// False
        /// </summary>
        off,
        /// <summary>
        /// True
        /// </summary>
        on,
        ///// <summary>
        ///// False
        ///// </summary>
        //[XmlEnum("0")]
        //Value0 = 0,

        ///// <summary>
        ///// True
        ///// </summary>
        //[XmlEnum("1")]
        //Value1 = 1,

        /// <summary>
        /// True
        /// </summary>
        [XmlEnum("true")]
        True,

        /// <summary>
        /// False
        /// </summary>
        [XmlEnum("false")]
        False,
    }
    /// <summary>
    /// Long Hexadecimal Number
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_LongHexNumber
    {
        public static CT_LongHexNumber Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_LongHexNumber ctObj = new CT_LongHexNumber();
            ctObj.val = XmlHelper.ReadBytes(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write("/>");
        }


        private byte[] valField;
        /// <summary>
        /// Four Digit Hexadecimal Number Value
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] val
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

    /// <summary>
    /// Two Digit Hexadecimal Number
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_ShortHexNumber
    {
        private byte[] valField;
        public static CT_ShortHexNumber Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ShortHexNumber ctObj = new CT_ShortHexNumber();
            ctObj.val = XmlHelper.ReadBytes(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        /// <summary>
        /// Two Digit Hexadecimal Number Value
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] val
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

    /// <summary>
    /// Two Digit Hexadecimal Number
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_UcharHexNumber
    {
        private byte[] valField;

        /// <summary>
        /// Two Digit Hexadecimal Number Value
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "hexBinary")]
        public byte[] val
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

    /// <summary>
    /// Decimal Number Value
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_DecimalNumber
    {
        public static CT_DecimalNumber Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DecimalNumber ctObj = new CT_DecimalNumber();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val, true);
            sw.Write("/>");
        }


        private string valField;

        /// <summary>
        /// Decimal Number
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
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

    /// <summary>
    /// Measurement in Twentieths of a Point
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_TwipsMeasure
    {

        private ulong valField;
        public static CT_TwipsMeasure Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TwipsMeasure ctObj = new CT_TwipsMeasure();
            ctObj.val = XmlHelper.ReadULong(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write("/>");
        }

        /// <summary>
        /// Measurement in Twentieths of a Point
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong val
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

    /// <summary>
    /// Signed Measurement in Twentieths of a Point
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SignedTwipsMeasure
    {

        public static CT_SignedTwipsMeasure Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SignedTwipsMeasure ctObj = new CT_SignedTwipsMeasure();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }


        private string valField;
        /// <summary>
        /// Signed Measurement in Twentieths of a Point
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
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

    /// <summary>
    /// Measurement in Pixels
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_PixelsMeasure
    {

        private ulong valField;

        /// <summary>
        /// Measurement in Pixels
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong val
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

    /// <summary>
    /// Half Point Measurement
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_HpsMeasure
    {
        public static CT_HpsMeasure Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_HpsMeasure ctObj = new CT_HpsMeasure();
            ctObj.val = XmlHelper.ReadULong(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val, true);
            sw.Write("/>");
        }


        private ulong valField;

        /// <summary>
        /// Half Point Measurement
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
        public ulong val
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
    /// <summary>
    /// Signed Measurement in Half-Points
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_SignedHpsMeasure
    {
        public static CT_SignedHpsMeasure Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SignedHpsMeasure ctObj = new CT_SignedHpsMeasure();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }
        private string valField;
        /// <summary>
        /// Signed Measurement in Half-Points
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "integer")]
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

    /// <summary>
    /// Name of Script Function
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_MacroName
    {

        private string valField;
        /// <summary>
        /// Script Subroutine Name Value
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

        public static CT_MacroName Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MacroName ctObj = new CT_MacroName();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val, true);
            sw.Write("/>");
        }
    }

    /// <summary>
    /// String
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_String
    {
        public static CT_String Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_String ctObj = new CT_String();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write("/>");
        }


        private string valField;
        /// <summary>
        /// String Value
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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

    /// <summary>
    /// Language Reference
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Lang
    {
        public static CT_Lang Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Lang ctObj = new CT_Lang();
            ctObj.val = XmlHelper.ReadString(node.Attributes["w:val"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<w:{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "w:val", this.val);
            sw.Write(">");
            sw.Write(string.Format("</w:{0}>", nodeName));
        }

        private string valField;
        /// <summary>
        /// Language Code
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified)]
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


    //<xsd:pattern value="\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}" />
    /// <summary>
    /// 128-Bit GUID
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", IsNullable = true)]
    public class CT_Guid
    {

        private string valField;
        /// <summary>
        /// GUID Value
        /// </summary>
        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, DataType = "token")]
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
}
