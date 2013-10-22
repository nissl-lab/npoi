using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Cell
    {

        private CT_CellFormula fField = null;

        private string vField = null;

        private CT_Rst isField = null;

        private CT_ExtensionList extLstField = null;

        private string rField = null;

        private uint? sField = null;

        private ST_CellType? tField = null;

        private uint? cmField = null;

        private uint? vmField = null;

        private bool? phField = null;

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
            fField = cell.fField;
            vField = cell.vField;
            isField = cell.isField;
            extLstField = cell.extLstField;
            rField = cell.rField;
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
            return fField != null;
        }
        public bool IsSetV()
        {
            return vField != null;
        }
        public bool IsSetIs()
        {
            return isField != null;
        }

        public void unsetF()
        {
            this.fField = null;
        }

        public void unsetV()
        {
            this.vField = null;
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
            this.isField = null;
        }

        [XmlElement]
        public CT_CellFormula f
        {
            get
            {
                return this.fField;
            }
            set
            {
                this.fField = value;
            }
        }
        [XmlElement]
        public string v
        {
            get
            {
                return this.vField;
            }
            set
            {
                this.vField = value;
            }
        }

        [XmlElement("is")]
        public CT_Rst @is
        {
            get
            {
                return this.isField;
            }
            set
            {
                this.isField = value;
            }
        }
        [XmlElement]
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
        [XmlAttribute]
        public string r
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
        [DefaultValue(typeof(uint), "0")]
        public uint s
        {
            get
            {
                return null == sField ? 0 : (uint)this.sField;
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
                return null == tField ? ST_CellType.n : (ST_CellType)this.tField;
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
                return null == cmField ? 0 : (uint)this.cmField;
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
                return null == vmField ? 0 : (uint)this.vmField;
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
                return null == phField ? false : (bool)this.phField;
            }
            set
            {
                this.phField = value;
            }
        }
    }
}
