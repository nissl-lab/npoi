using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "calcChain")]
    public class CT_CalcChain
    {

        private List<CT_CalcCell> cField = new List<CT_CalcCell>(); // [1..*]    Cell
        private CT_ExtensionList extLstField = null; //  [0..1]  

        public int SizeOfCArray()
        {
            return this.c.Count;
        }

        public CT_CalcCell GetCArray(int index)
        {
            return c[index];
        }
        public void AddC(CT_CalcCell cell)
        {
            this.c.Add(cell);
        }
        public void RemoveC(int index)
        {
            this.c.RemoveAt(index);
        }
        [XmlElement("c")]
        public List<CT_CalcCell> c
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
        [XmlIgnore]
        public bool extLstSpecified
        {
            get { return null != this.extLst; }
        }
    }
}
