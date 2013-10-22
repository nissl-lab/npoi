using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Row
    {

        private List<CT_Cell> cField = null; // optional element

        private CT_ExtensionList extLstField = null; // optional element

        // the following are all optional attributes
        private uint? rField = null;

        private string spansField = null; // a region is contained in this field, e.g. "1:3"

        private uint? sField = null;

        private bool? customFormatField = null;

        private double? htField = null;

        private bool? hiddenField = null;

        private bool? customHeightField = null;

        private byte? outlineLevelField = null;

        private bool? collapsedField = null;

        private bool? thickTopField = null;

        private bool? thickBotField = null;

        private bool? phField = null;

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
        public void unsetCollapsed()
        {
            this.collapsedField = null;
        }
        public void unSetS()
        {
            this.sField = null;
        }
        public void unSetCustomFormat()
        {
            this.customFormatField = null;
        }
        public bool IsSetHidden()
        {
            return this.hiddenField != null;
        }

        public bool IsSetCollapsed()
        {
            return this.collapsedField != null;
        }
        public bool IsSetHt()
        {
            return this.htField != null;
        }
        public void unSetHt()
        {
            this.htField = null;
        }
        public bool IsSetCustomHeight()
        {
            return this.customHeightField != null;
        }
        public void unSetCustomHeight()
        {
            this.customHeightField = null;
        }
        public bool IsSetS()
        {
            return this.sField != null;
        }
        public void unsetHidden()
        {
            this.hiddenField = null;
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
        [XmlIgnore]
        public bool cSpecified
        {
            get { return null != this.cField; }
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
            get { return null != this.extLstField; }
        }

        [XmlAttribute("r")]
        public uint r
        {
            get
            {
                return null == this.rField ? 0 : (uint)this.rField;
            }
            set
            {
                this.rField = value;
            }
        }
        [XmlIgnore]
        public bool rSpecified
        {
            get { return null != this.rField; }
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
        [XmlIgnore]
        public bool spansSpecified
        {
            get { return (null != spansField); }
        }

        //[DefaultValue(typeof(uint), "0")]
        [XmlAttribute]
        public uint s
        {
            get
            {
                return (null == sField) ? 0 : (uint)this.sField;
            }
            set
            {
                this.sField = value;
            }
        }
        [XmlIgnore]
        public bool sSpecified
        {
            get { return (null != sField); }
            set { CT_Row.SetSpecifiedWithDefaultValue(sField, value); }
        }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool customFormat
        {
            get
            {
                return (null == customFormatField) ? false : (bool)this.customFormatField;
            }
            set
            {
                this.customFormatField = value;
            }
        }
        [XmlIgnore]
        public bool customFormatSpecified
        {
            get { return (null != customFormatField); }
            set { CT_Row.SetSpecifiedWithDefaultFalse(customFormatField, value); }
        }

        [XmlAttribute]
        public double ht
        {
            get
            {
                return (double)this.htField;
            }
            set
            {
                this.htField = value;
            }
        }

        [XmlIgnore]
        public bool htSpecified
        {
            get { return null != this.htField; }
        }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool hidden
        {
            get
            {
                return (null == hiddenField) ? false : (bool)this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }
        [XmlIgnore]
        public bool hiddenSpecified
        {
            get { return (null != hiddenField); }
            set { CT_Row.SetSpecifiedWithDefaultFalse(hiddenField, value); }
        }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool customHeight
        {
            get
            {
                return (null == customHeightField) ? false : (bool)this.customHeightField;
            }
            set
            {
                this.customHeightField = value;
            }
        }
        [XmlIgnore]
        public bool customHeightSpecified
        {
            get { return (null != customHeightField); }
            set { CT_Row.SetSpecifiedWithDefaultFalse(customHeightField, value); }
        }

        [DefaultValue(typeof(byte), "0")]
        [XmlAttribute]
        public byte outlineLevel
        {
            get
            {
                return (null == outlineLevelField) ? (byte)0 : (byte)this.outlineLevelField;
            }
            set
            {
                this.outlineLevelField = value;
            }
        }
        [XmlIgnore]
        public bool outlineLevelSpecified
        {
            get { return (null != outlineLevelField); }
            set { CT_Row.SetSpecifiedWithDefaultValue(outlineLevelField, value); }
        }

        //[DefaultValue(false)]
        [XmlAttribute]
        public bool collapsed
        {
            get
            {
                return (null == collapsedField) ? false : (bool)this.collapsedField;
            }
            set
            {
                this.collapsedField = value;
            }
        }
        [XmlIgnore]
        public bool collapsedSpecified
        {
            get { return (null != collapsedField); }
            set { CT_Row.SetSpecifiedWithDefaultFalse(collapsedField, value); }
        }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool thickTop
        {
            get
            {
                return (null == thickTopField) ? false : (bool)this.thickTopField;
            }
            set
            {
                this.thickTopField = value;
            }
        }
        [XmlIgnore]
        public bool thickTopSpecified
        {
            get { return (null != thickTopField); }
            set { CT_Row.SetSpecifiedWithDefaultFalse(thickTopField, value); }
        }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool thickBot
        {
            get
            {
                return (null == thickBotField) ? false : (bool)this.thickBotField;
            }
            set
            {
                this.thickBotField = value;
            }
        }
        [XmlIgnore]
        public bool thickBotSpecified
        {
            get { return (null != thickBotField); }
            set { CT_Row.SetSpecifiedWithDefaultFalse(thickBotField, value); }
        }

        [DefaultValue(false)]
        [XmlAttribute]
        public bool ph
        {
            get
            {
                return (null == phField) ? false : (bool)this.phField;
            }
            set
            {
                this.phField = value;
            }
        }
        [XmlIgnore]
        public bool phSpecified
        {
            get { return (null != phField); }
            set { CT_Row.SetSpecifiedWithDefaultFalse(phField, value); }
        }

        /// <summary>
        /// Set the Nullable bool to reflect Specified or not.
        /// Preserves bool value true. 
        /// </summary>
        /// <param name="field">field to set to specified or not</param>
        /// <param name="specified">true or false</param>
        public static void SetSpecifiedWithDefaultFalse(bool? field, bool specified)
        {
            if (specified)
            {
                field = field.HasValue ? field.Value : false; // preserve existing bool value, default to false
            }
            else
            {
                field = null;
            }
        }
        /// <summary>
        /// Set the Nullable value type to reflect Specified or not.
        /// Preserves set values. 
        /// </summary>
        /// <param name="field">field to set to specified or not</param>
        /// <param name="specified">true or false</param>
        public static void SetSpecifiedWithDefaultValue<T>(T? field, bool specified) where T : struct
        {
            if (specified)
            {
                field = field.HasValue ? field.Value : default(T);
            }
            else
            {
                field = null;
            }
        }
    }

}
