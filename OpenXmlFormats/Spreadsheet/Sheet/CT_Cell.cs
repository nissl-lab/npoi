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
    public class CT_Cell
    {
        private uint? sField = null;

        private ST_CellType? tField = null;

        private uint? cmField = null;

        private uint? vmField = null;

        private bool? phField = null;

        public static CT_Cell Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Cell ctObj = new CT_Cell();
            ctObj.r = XmlHelper.ReadString(node.Attributes[nameof(r)]);
            ctObj.s = XmlHelper.ReadUInt(node.Attributes[nameof(s)]);
            if (node.Attributes[nameof(t)] != null)
                ctObj.t = (ST_CellType)Enum.Parse(typeof(ST_CellType), node.Attributes[nameof(t)].Value);
            ctObj.cm = XmlHelper.ReadUInt(node.Attributes[nameof(cm)]);
            ctObj.vm = XmlHelper.ReadUInt(node.Attributes[nameof(vm)]);
            ctObj.ph = XmlHelper.ReadBool(node.Attributes[nameof(ph)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(f))
                    ctObj.f = CT_CellFormula.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(v))
                    ctObj.v = childNode.InnerText;
                else if (childNode.LocalName == nameof(@is))
                    ctObj.@is = CT_Rst.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");

            XmlHelper.WriteAttribute(sw, nameof(r), this.r);
            XmlHelper.WriteAttribute(sw, nameof(s), this.s);
            if (this.t != ST_CellType.n)
                XmlHelper.WriteAttribute(sw, nameof(t), this.t.ToString());
            XmlHelper.WriteAttribute(sw, nameof(cm), this.cm);
            XmlHelper.WriteAttribute(sw, nameof(vm), this.vm);
            XmlHelper.WriteAttribute(sw, nameof(ph), this.ph, false);
            if (this.f == null
                && this.v == null
                && this.@is == null
                && this.extLst == null)
            {
                sw.Write("/>");
            }
            else
            {
                // changed to null conditional
                sw.Write(">");
                this.f?.Write(sw, nameof(f));
                if (this.v != null)
                    sw.Write($"<v>{XmlHelper.EncodeXml(this.v)}</v>");
                this.@is?.Write(sw, nameof(@is));
                this.extLst?.Write(sw, nameof(extLst));
                sw.Write($"</{nodeName}>");
            }
        }


        //public CT_Cell()
        //{
        //    this.extLstField = new CT_ExtensionList();
        //    //this.isField = new CT_Rst();
        //    //this.fField = new CT_CellFormula();
        //    this.sField = (uint)(0);
        //    this.tField = ST_CellType.n;
        //    this.cmField = ((uint)(0));
        //    this.vmField = ((uint)(0));
        //    this.phField = false;
        //}
        public void Set(CT_Cell cell)
        {
            f = cell.f;
            v = cell.v;
            @is = cell.@is;
            extLst = cell.extLst;
            r = cell.r;
            sField = cell.sField;
            tField = cell.tField;
            cmField = cell.cmField;
            vmField = cell.vmField;
            phField = cell.phField;
        }
        public bool IsSetT()
        {
            return tField != ST_CellType.n;
        }
        public bool IsSetS()
        {
            return sField != 0;
        }
        public bool IsSetF()
        {
            return f != null;
        }
        public bool IsSetV()
        {
            return v != null;
        }
        public bool IsSetIs()
        {
            return @is != null;
        }
        public bool IsSetR()
        {
            return r != null;
        }
        public void unsetF()
        {
            this.f = null;
        }

        public void unsetV()
        {
            this.v = null;
        }

        public void unsetS()
        {
            this.sField = 0;
        }
        public void unsetT()
        {
            this.tField = ST_CellType.n;
        }

        public void unsetIs()
        {
            this.@is = null;
        }
        public void unsetR()
        {
            this.r = null;
        }
        [XmlElement]
        public CT_CellFormula f { get; set; } = null;

        [XmlElement]
        public string v { get; set; } = null;

        [XmlElement(nameof(@is))]
        public CT_Rst @is { get; set; } = null;

        [XmlElement]
        public CT_ExtensionList extLst { get; set; } = null;

        [XmlAttribute]
        public string r { get; set; } = null;

        [XmlAttribute]
        [DefaultValue(typeof(uint), "0")]
        public uint s
        {
            get
            {
                return this.sField ?? 0;
            }
            set
            {
                this.sField = value;
            }
        }
        [XmlAttribute]
        [DefaultValue(ST_CellType.n)]
        public ST_CellType t
        {
            get
            {
                return this.tField ?? ST_CellType.n;
            }
            set
            {
                this.tField = value;
            }
        }
        [XmlAttribute]
        [DefaultValue(typeof(uint), "0")]
        public uint cm
        {
            get
            {
                return this.cmField ?? 0;
            }
            set
            {
                this.cmField = value;
            }
        }
        [XmlAttribute]
        [DefaultValue(typeof(uint), "0")]
        public uint vm
        {
            get
            {
                return this.vmField ?? 0;
            }
            set
            {
                this.vmField = value;
            }
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool ph
        {
            get
            {
                return this.phField ?? false;
            }
            set
            {
                this.phField = value;
            }
        }
    }
}
