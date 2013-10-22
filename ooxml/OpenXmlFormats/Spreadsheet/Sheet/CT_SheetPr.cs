using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_SheetPr
    {

        private CT_Color tabColorField;

        private CT_OutlinePr outlinePrField;

        private CT_PageSetUpPr pageSetUpPrField;

        private bool syncHorizontalField;

        private bool syncVerticalField;

        private string syncRefField;

        private bool transitionEvaluationField;

        private bool transitionEntryField;

        private bool publishedField;

        private string codeNameField;

        private bool filterModeField;

        private bool enableFormatConditionsCalculationField;

        public CT_SheetPr()
        {
            //this.pageSetUpPrField = new CT_PageSetUpPr();
            //this.outlinePrField = new CT_OutlinePr();
            this.tabColorField = new CT_Color();
            this.syncHorizontalField = false;
            this.syncVerticalField = false;
            this.transitionEvaluationField = false;
            this.transitionEntryField = false;
            this.publishedField = true;
            this.filterModeField = false;
            this.enableFormatConditionsCalculationField = true;
        }
        public bool IsSetOutlinePr()
        {
            return this.outlinePrField != null;
        }
        public bool IsSetPageSetUpPr()
        {
            return this.pageSetUpPrField != null;
        }
        public CT_PageSetUpPr AddNewPageSetUpPr()
        {
            this.pageSetUpPrField = new CT_PageSetUpPr();
            return this.pageSetUpPrField;
        }
        public CT_OutlinePr AddNewOutlinePr()
        {
            this.outlinePrField = new CT_OutlinePr();
            return this.outlinePrField;
        }


        public CT_Color tabColor
        {
            get
            {
                return this.tabColorField;
            }
            set
            {
                this.tabColorField = value;
            }
        }

        public CT_OutlinePr outlinePr
        {
            get
            {
                return this.outlinePrField;
            }
            set
            {
                this.outlinePrField = value;
            }
        }

        public CT_PageSetUpPr pageSetUpPr
        {
            get
            {
                return this.pageSetUpPrField;
            }
            set
            {
                this.pageSetUpPrField = value;
            }
        }

        [DefaultValue(false)]
        public bool syncHorizontal
        {
            get
            {
                return this.syncHorizontalField;
            }
            set
            {
                this.syncHorizontalField = value;
            }
        }

        [DefaultValue(false)]
        public bool syncVertical
        {
            get
            {
                return this.syncVerticalField;
            }
            set
            {
                this.syncVerticalField = value;
            }
        }

        public string syncRef
        {
            get
            {
                return this.syncRefField;
            }
            set
            {
                this.syncRefField = value;
            }
        }

        [DefaultValue(false)]
        public bool transitionEvaluation
        {
            get
            {
                return this.transitionEvaluationField;
            }
            set
            {
                this.transitionEvaluationField = value;
            }
        }

        [DefaultValue(false)]
        public bool transitionEntry
        {
            get
            {
                return this.transitionEntryField;
            }
            set
            {
                this.transitionEntryField = value;
            }
        }

        [DefaultValue(true)]
        public bool published
        {
            get
            {
                return this.publishedField;
            }
            set
            {
                this.publishedField = value;
            }
        }

        public string codeName
        {
            get
            {
                return this.codeNameField;
            }
            set
            {
                this.codeNameField = value;
            }
        }

        [DefaultValue(false)]
        public bool filterMode
        {
            get
            {
                return this.filterModeField;
            }
            set
            {
                this.filterModeField = value;
            }
        }

        [DefaultValue(true)]
        public bool enableFormatConditionsCalculation
        {
            get
            {
                return this.enableFormatConditionsCalculationField;
            }
            set
            {
                this.enableFormatConditionsCalculationField = value;
            }
        }

        public bool IsSetTabColor()
        {
            return this.tabColor != null;
        }
    }

}
