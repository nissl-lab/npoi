using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml
{

    public interface EG_ObjectChoices
    { 
        
    }
    public class CT_ShapeNonVisual
    {
        private CT_NonVisualDrawingProps cNvPrField;
        private CT_NonVisualDrawingShapeProps cNvSpPrField;

        public CT_NonVisualDrawingProps cNvPr
        {
            get
            {
                return this.cNvPrField;
            }
            set
            {
                this.cNvPrField = value;
            }
        }

        public CT_NonVisualDrawingShapeProps cNvSpPr
        {
            get
            {
                return this.cNvSpPrField;
            }
            set
            {
                this.cNvSpPrField = value;
            }
        }
    }
    public class CT_Shape : EG_ObjectChoices
    {
        private CT_ShapeNonVisual nvSpPrField;
        private CT_ShapeProperties spPrField;
        private CT_ShapeStyle styleField;
        private CT_TextBody txBodyField;

        private string macroField;
        private string textlinkField;
        private bool fLocksTextField;
        private bool fPublishedField;

        public CT_ShapeNonVisual nvSpPr
        {
            get
            {
                return this.nvSpPrField;
            }
            set
            {
                this.nvSpPrField = value;
            }
        }
        public CT_ShapeProperties spPr
        {
            get
            {
                return this.spPrField;
            }
            set
            {
                this.spPrField = value;
            }
        }
        public CT_TextBody txBody
        {
            get
            {
                return this.txBodyField;
            }
            set
            {
                this.txBodyField = value;
            }
        }
        public CT_ShapeStyle style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }

        [XmlAttributeAttribute()]
        public string macro
        {
            get { return macroField; }
            set { macroField = value; }
        }

        [XmlAttributeAttribute()]
        public string textlink
        {
            get { return textlinkField; }
            set { textlinkField = value; }
        }
        [XmlAttributeAttribute()]
        public bool fLocksText
        {
            get { return fLocksTextField; }
            set { fLocksTextField = value; }
        }

        [XmlAttributeAttribute()]
        public bool fPublished
        {
            get { return fPublishedField; }
            set { fPublishedField = value; }
        }
    }
    public class CT_ConnectorNonVisual
    {
        private CT_NonVisualDrawingProps cNvPrField;
        private CT_NonVisualConnectorProperties cNvCxnSpPrField;

        public CT_NonVisualConnectorProperties cNvCxnSpPr
        {
            get
            {
                return this.cNvCxnSpPrField;
            }
            set
            {
                this.cNvCxnSpPrField = value;
            }
        }
        public CT_NonVisualDrawingProps cNvPr
        {
            get
            {
                return this.cNvPrField;
            }
            set
            {
                this.cNvPrField = value;
            }
        }
    }
    public class CT_Marker
    {
        public int col
        {
            get;set;
        }
        public long colOff
        {
        get;set;
        }
        public int row
        {
        get;set;
        }
        public long rowOff
        {
        get;set;
        }
    }
    public enum ST_EditAs
    {
        twoCell,
        oneCell,
        absolute
    }
    public class CT_AnchorClientData
    {
        [XmlAttribute]
        public bool fLocksWithSheet
        {
            get;
            set;
        }
        [XmlAttribute]
        public bool fPrintsWithSheet
        {
            get;
            set;
        }
    }
    public class CT_TwoCellAnchor
    {
        private CT_Marker fromField;
        private CT_Marker toField;
        private CT_AnchorClientData clientDataField;

        public CT_Marker from
        {
            get { return fromField; }
            set { fromField = value; }
        }
        public CT_Marker to
        {
            get { return toField; }
            set { toField = value; }
        }
        public ST_EditAs editAs
        {
            get;
            set;
        }
        public CT_AnchorClientData clientData
        {
            get { return clientDataField; }
            set { clientDataField = value; } 
        }
    }
    public class CT_Connector : EG_ObjectChoices
    { 
        string macroField;
        bool fPublishField;
        private CT_ShapeProperties spPrField;
        private CT_ShapeStyle styleField;
        private CT_ConnectorNonVisual nvCxnSpPrField;
        public CT_ConnectorNonVisual nvCxnSpPr
        {
            get { return nvCxnSpPrField; }
            set { nvCxnSpPrField = value; }
        }
        public CT_ShapeProperties spPr
        {
            get
            {
                return this.spPrField;
            }
            set
            {
                this.spPrField = value;
            }
        }
         public CT_ShapeStyle style
         {
             get
             {
                 return this.styleField;
             }
             set
             {
                 this.styleField = value;
             }
         }
         [XmlAttributeAttribute()]
        public string macro
        {
            get{return this.macroField;}
            set{this.macroField=value;}
        }
        [XmlAttributeAttribute()]
        public bool fPublished
        {
            get{return this.fPublishField;}
            set{this.fPublishField = value;}
        }
    }
    public partial class CT_Picture:EG_ObjectChoices
    { 
        private CT_ShapeStyle styleField;
        private string macroField;
        private bool fPublishedField;

        [XmlAttributeAttribute()]
        public string macro
        {
            get { return macroField; }
            set { macroField = value; }
        }

        [XmlAttributeAttribute()]
        public bool fPublished
        {
            get { return fPublishedField; }
            set { fPublishedField = value; }
        }
        public CT_ShapeStyle style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }
    }
}
