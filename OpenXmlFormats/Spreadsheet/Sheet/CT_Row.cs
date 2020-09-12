using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Row
    {

        private List<CT_Cell> cField = null; // optional element

        private CT_ExtensionList extLstField = null; // optional element

        // the following are all optional attributes
        private uint rField;

        private string spansField = null; // a region is contained in this field, e.g. "1:3"

        private uint sField;

        private bool customFormatField;

        private double htField=-1;

        private bool hiddenField;

        private bool customHeightField;

        private byte outlineLevelField;

        private bool collapsedField;

        private bool thickTopField;

        private bool thickBotField;

        private bool phField;
        private double dyDescentField; //x14ac:dyDescent

        public static CT_Row Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Row ctObj = new CT_Row();
            ctObj.r = XmlHelper.ReadUInt(node.Attributes["r"]);
            ctObj.spans = XmlHelper.ReadString(node.Attributes["spans"]);
            ctObj.s = XmlHelper.ReadUInt(node.Attributes["s"]);
            ctObj.customFormat = XmlHelper.ReadBool(node.Attributes["customFormat"]);
            ctObj.dyDescentField = XmlHelper.ReadDouble(node.Attributes["x14ac:dyDescent"]);
            if (node.Attributes["ht"]!=null)
                ctObj.ht = XmlHelper.ReadDouble(node.Attributes["ht"]);
            ctObj.hidden = XmlHelper.ReadBool(node.Attributes["hidden"]);
            ctObj.outlineLevel = XmlHelper.ReadByte(node.Attributes["outlineLevel"]);
            ctObj.customHeight = XmlHelper.ReadBool(node.Attributes["customHeight"]);
            ctObj.collapsed = XmlHelper.ReadBool(node.Attributes["collapsed"]);
            ctObj.thickTop = XmlHelper.ReadBool(node.Attributes["thickTop"]);
            ctObj.thickBot = XmlHelper.ReadBool(node.Attributes["thickBot"]);
            ctObj.ph = XmlHelper.ReadBool(node.Attributes["ph"]);
            ctObj.c = new List<CT_Cell>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "c")
                    ctObj.c.Add(CT_Cell.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }

        public bool IsSetR()
        {
            return this.rField > 0;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "r", this.r);
            XmlHelper.WriteAttribute(sw, "spans", this.spans);
            XmlHelper.WriteAttribute(sw, "s", this.s);
            XmlHelper.WriteAttribute(sw, "customFormat", this.customFormat,false);
            if (this.ht>=0)
                XmlHelper.WriteAttribute(sw, "ht", this.ht);
            XmlHelper.WriteAttribute(sw, "hidden", this.hidden,false);
            XmlHelper.WriteAttribute(sw, "customHeight", this.customHeight,false);
            XmlHelper.WriteAttribute(sw, "outlineLevel", this.outlineLevel);
            XmlHelper.WriteAttribute(sw, "collapsed", this.collapsed, false);
            XmlHelper.WriteAttribute(sw, "thickTop", this.thickTop,false);
            XmlHelper.WriteAttribute(sw, "thickBot", this.thickBot,false);
            XmlHelper.WriteAttribute(sw, "ph", this.ph, false);
            XmlHelper.WriteAttribute(sw, "x14ac:dyDescent", this.dyDescentField, false);
            sw.Write(">");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
            if (this.c != null)
            {
                foreach (CT_Cell x in this.c)
                {
                    x.Write(sw, "c");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }



        public void Set(CT_Row row)
        {
            cField = row.cField;
            extLstField = row.extLstField;
            rField = row.rField;
            spansField = row.spansField;
            sField = row.sField;
            customFormatField = row.customFormatField;
            htField = row.htField;
            hiddenField = row.hiddenField;
            customHeightField = row.customHeightField;
            outlineLevelField = row.outlineLevelField;
            collapsedField = row.collapsedField;
            thickTopField = row.thickTopField;
            thickBotField = row.thickBotField;
            phField = row.phField;
        }
        public CT_Cell AddNewC()
        {
            if (null == cField) { cField = new List<CT_Cell>(); }
            CT_Cell cell = new CT_Cell();
            this.cField.Add(cell);
            return cell;
        }
        public void UnsetCollapsed()
        {
            this.collapsedField = false;
        }
        public void UnsetS()
        {
            this.sField = 0;
        }
        public void UnsetCustomFormat()
        {
            this.customFormatField = false;
        }
        public bool IsSetHidden()
        {
            return this.hiddenField != false;
        }
        public bool IsSetCollapsed()
        {
            return this.collapsedField != false;
        }
        public bool IsSetHt()
        {
            return this.htField >=0;
        }
        public void unSetHt()
        {
            this.htField = -1;
        }
        public bool IsSetCustomHeight()
        {
            return this.customHeightField != false;
        }
        public void unSetCustomHeight()
        {
            this.customHeightField = false;
        }
        public bool IsSetS()
        {
            return this.sField != 0;
        }
        public void unsetHidden()
        {
            this.hiddenField = false;
        }

        public int SizeOfCArray()
        {
            return (null == cField) ? 0 : cField.Count;
        }
        public CT_Cell GetCArray(int index)
        {
            return (null == cField) ? null : cField[index];
        }
        public void SetCArray(CT_Cell[] array)
        {
            cField = new List<CT_Cell>(array);
        }
        [XmlElement("c")]
        public List<CT_Cell> c
        {
            get
            {
                return this.cField;
            }
            set
            {
                this.cField = value;
            }
        }
        [XmlElement("extLst")]
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

        [XmlAttribute("r")]
        public uint r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }

        [XmlAttribute]
        public string spans
        {
            get
            {
                return this.spansField;
            }
            set
            {
                this.spansField = value;
            }
        }

        //[DefaultValue(typeof(uint), "0")]
        [XmlAttribute]
        public uint s
        {
            get
            {
                return this.sField;
            }
            set
            {
                this.sField = value;
            }
        }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool customFormat
        {
            get
            {
                return this.customFormatField;
            }
            set
            {
                this.customFormatField = value;
            }
        }
        [XmlAttribute]
        public double ht
        {
            get
            {
                return this.htField;
            }
            set
            {
                this.htField = value;
            }
        }


        //[DefaultValue(false)]
        [XmlAttribute]
        public bool hidden
        {
            get
            {
                return this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool customHeight
        {
            get
            {
                return this.customHeightField;
            }
            set
            {
                this.customHeightField = value;
            }
        }

        [DefaultValue(typeof(byte), "0")]
        [XmlAttribute]
        public byte outlineLevel
        {
            get
            {
                return this.outlineLevelField;
            }
            set
            {
                this.outlineLevelField = value;
            }
        }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool collapsed
        {
            get
            {
                return this.collapsedField;
            }
            set
            {
                this.collapsedField = value;
            }
        }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool thickTop
        {
            get
            {
                return this.thickTopField;
            }
            set
            {
                this.thickTopField = value;
            }
        }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool thickBot
        {
            get
            {
                return this.thickBotField;
            }
            set
            {
                this.thickBotField = value;
            }
        }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool ph
        {
            get
            {
                return this.phField;
            }
            set
            {
                this.phField = value;
            }
        }
    }

}
