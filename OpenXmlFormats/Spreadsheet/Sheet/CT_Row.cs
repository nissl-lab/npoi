using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
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
        public static CT_Row Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Row ctObj = new CT_Row
            {
                r = XmlHelper.ReadUInt(node.Attributes[nameof(r)]),
                spans = XmlHelper.ReadString(node.Attributes[nameof(spans)]),
                s = XmlHelper.ReadUInt(node.Attributes[nameof(s)]),
                customFormat = XmlHelper.ReadBool(node.Attributes[nameof(customFormat)])
            };
            if (node.Attributes[nameof(ht)] != null)
                ctObj.ht = XmlHelper.ReadDouble(node.Attributes[nameof(ht)]);
            ctObj.hidden = XmlHelper.ReadBool(node.Attributes[nameof(hidden)]);
            ctObj.outlineLevel = XmlHelper.ReadByte(node.Attributes[nameof(outlineLevel)]);
            ctObj.customHeight = XmlHelper.ReadBool(node.Attributes[nameof(customHeight)]);
            ctObj.collapsed = XmlHelper.ReadBool(node.Attributes[nameof(collapsed)]);
            ctObj.thickTop = XmlHelper.ReadBool(node.Attributes[nameof(thickTop)]);
            ctObj.thickBot = XmlHelper.ReadBool(node.Attributes[nameof(thickBot)]);
            ctObj.ph = XmlHelper.ReadBool(node.Attributes[nameof(ph)]);
            ctObj.c = new List<CT_Cell>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(c))
                    ctObj.c.Add(CT_Cell.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(r), this.r);
            XmlHelper.WriteAttribute(sw, nameof(spans), this.spans);
            XmlHelper.WriteAttribute(sw, nameof(s), this.s);
            XmlHelper.WriteAttribute(sw, nameof(customFormat), this.customFormat, false);
            if (this.ht >= 0)
                XmlHelper.WriteAttribute(sw, nameof(ht), this.ht);
            XmlHelper.WriteAttribute(sw, nameof(hidden), this.hidden, false);
            XmlHelper.WriteAttribute(sw, nameof(customHeight), this.customHeight, false);
            XmlHelper.WriteAttribute(sw, nameof(outlineLevel), this.outlineLevel);
            XmlHelper.WriteAttribute(sw, nameof(collapsed), this.collapsed, false);
            XmlHelper.WriteAttribute(sw, nameof(thickTop), this.thickTop, false);
            XmlHelper.WriteAttribute(sw, nameof(thickBot), this.thickBot, false);
            XmlHelper.WriteAttribute(sw, nameof(ph), this.ph, false);
            sw.Write(">");
            this.extLst?.Write(sw, nameof(extLst));
            this.c?.ForEach(x => x.Write(sw, nameof(c)));
            sw.Write($"</{nodeName}>");
        }



        public void Set(CT_Row row)
        {
            c = row.c;
            extLst = row.extLst;
            r = row.r;
            spans = row.spans;
            s = row.s;
            customFormat = row.customFormat;
            ht = row.ht;
            hidden = row.hidden;
            customHeight = row.customHeight;
            outlineLevel = row.outlineLevel;
            collapsed = row.collapsed;
            thickTop = row.thickTop;
            thickBot = row.thickBot;
            ph = row.ph;
        }
        public CT_Cell AddNewC()
        {
            if (null == c) { c = new List<CT_Cell>(); }
            CT_Cell cell = new CT_Cell();
            this.c.Add(cell);
            return cell;
        }
        public void UnsetCollapsed()
        {
            this.collapsed = false;
        }
        public void UnsetS()
        {
            this.s = 0;
        }
        public void UnsetCustomFormat()
        {
            this.customFormat = false;
        }
        public bool IsSetHidden()
        {
            return this.hidden != false;
        }
        public bool IsSetCollapsed()
        {
            return this.collapsed != false;
        }
        public bool IsSetHt()
        {
            return this.ht >= 0;
        }
        public void unSetHt()
        {
            this.ht = -1;
        }
        public bool IsSetCustomHeight()
        {
            return this.customHeight != false;
        }
        public void unSetCustomHeight()
        {
            this.customHeight = false;
        }
        public bool IsSetS()
        {
            return this.s != 0;
        }
        public void unsetHidden()
        {
            this.hidden = false;
        }

        public int SizeOfCArray()
        {
            return c?.Count ?? 0;
        }
        public CT_Cell GetCArray(int index)
        {
            return c?[index];
        }
        public void SetCArray(IEnumerable<CT_Cell> cells)
        {
            c = cells.ToList();
        }
        [XmlElement(nameof(c))]
        public List<CT_Cell> c { get; set; } = null;
        [XmlElement(nameof(extLst))]
        public CT_ExtensionList extLst { get; set; } = null;

        [XmlAttribute(nameof(r))]
        public uint r { get; set; }

        [XmlAttribute]
        public string spans { get; set; } = null;

        //[DefaultValue(typeof(uint), "0")]
        [XmlAttribute]
        public uint s { get; set; }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool customFormat { get; set; }
        [XmlAttribute]
        public double ht { get; set; } = -1;


        //[DefaultValue(false)]
        [XmlAttribute]
        public bool hidden { get; set; }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool customHeight { get; set; }

        [DefaultValue(typeof(byte), "0")]
        [XmlAttribute]
        public byte outlineLevel { get; set; }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool collapsed { get; set; }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool thickTop { get; set; }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool thickBot { get; set; }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool ph { get; set; }
    }

}
