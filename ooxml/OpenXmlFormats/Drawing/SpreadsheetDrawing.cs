using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.Collections.Generic;
using System.ComponentModel;

namespace NPOI.OpenXmlFormats.Dml.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_ShapeNonVisual
    {
        private CT_NonVisualDrawingProps cNvPrField;
        private CT_NonVisualDrawingShapeProps cNvSpPrField;

        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPrField = new CT_NonVisualDrawingProps();
            return this.cNvPrField;
        }
        public CT_NonVisualDrawingShapeProps AddNewCNvSpPr()
        {
            this.cNvSpPrField = new CT_NonVisualDrawingShapeProps();
            return this.cNvSpPrField;
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

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_GraphicalObjectFrameNonVisual
    {

        CT_NonVisualDrawingProps cNvPrField;
        CT_NonVisualGraphicFrameProperties cNvGraphicFramePrField;

        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPrField = new CT_NonVisualDrawingProps();
            return this.cNvPrField;
        }
        public CT_NonVisualGraphicFrameProperties AddNewCNvGraphicFramePr()
        {
            this.cNvGraphicFramePrField = new CT_NonVisualGraphicFrameProperties();
            return this.cNvGraphicFramePrField;
        }

        public CT_NonVisualDrawingProps cNvPr
        {
            get { return cNvPrField; }
            set { cNvPrField = value; }
        }
        public CT_NonVisualGraphicFrameProperties cNvGraphicFramePr
        {
            get { return cNvGraphicFramePrField; }
            set { cNvGraphicFramePrField = value; }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_GraphicalObjectFrame
    {
        CT_GraphicalObjectFrameNonVisual nvGraphicFramePrField;
        CT_Transform2D xfrmField;
        CT_GraphicalObject graphicField;
        private string macroField;
        private bool fPublishedField;

        public void Set(CT_GraphicalObjectFrame obj)
        {
            this.xfrmField = obj.xfrmField;
            this.graphicField = obj.graphicField;
            this.nvGraphicFramePrField = obj.nvGraphicFramePrField;
            this.macroField = obj.macroField;
            this.fPublishedField = obj.fPublishedField;
        }

        public CT_Transform2D AddNewXfrm()
        {
            this.xfrmField = new CT_Transform2D();
            return this.xfrmField;
        }
        public CT_GraphicalObject AddNewGraphic()
        {
            this.graphicField = new CT_GraphicalObject();
            return this.graphicField;
        }

        public CT_GraphicalObjectFrameNonVisual AddNewNvGraphicFramePr()
        {
            this.nvGraphicFramePr = new CT_GraphicalObjectFrameNonVisual();
            return this.nvGraphicFramePr;
        }
        [XmlElement]
        public CT_GraphicalObjectFrameNonVisual nvGraphicFramePr
        {
            get { return nvGraphicFramePrField; }
            set { nvGraphicFramePrField = value; }
        }
        [XmlElement]
        public CT_Transform2D xfrm
        {
            get { return xfrmField; }
            set { xfrmField = value; }
        }
        [XmlAttribute]
        public string macro
        {
            get { return macroField; }
            set { macroField = value; }
        }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool fPublished
        {
            get { return fPublishedField; }
            set { fPublishedField = value; }
        }
        [XmlElement]
        public CT_GraphicalObject graphic
        {
            get { return graphicField; }
            set { graphicField = value; }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Shape // empty interface: EG_ObjectChoices
    {
        private CT_ShapeNonVisual nvSpPrField;
        private CT_ShapeProperties spPrField;
        private CT_ShapeStyle styleField;
        private CT_TextBody txBodyField;

        private string macroField;
        private string textlinkField;
        private bool fLocksTextField;
        private bool fPublishedField;

        public void Set(CT_Shape obj)
        {
            this.macroField = obj.macro;
            this.textlinkField = obj.textlink;
            this.fLocksTextField = obj.fLocksText;
            this.fPublishedField = obj.fPublished;

            this.nvSpPrField = obj.nvSpPr;
            this.spPrField = obj.spPr;
            this.styleField = obj.style;
            this.txBodyField = obj.txBody;
        }

        public CT_ShapeNonVisual AddNewNvSpPr()
        {
            this.nvSpPrField = new CT_ShapeNonVisual();
            return this.nvSpPrField;
        }

        public CT_ShapeProperties AddNewSpPr()
        {
            this.spPrField = new CT_ShapeProperties();
            return this.spPrField;
        }
        public CT_ShapeStyle AddNewStyle()
        {
            this.styleField = new CT_ShapeStyle();
            return this.styleField;
        }
        public CT_TextBody AddNewTxBody()
        {
            this.txBodyField = new CT_TextBody();
            return this.txBodyField;
        }

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

        [XmlAttribute]
        public string macro
        {
            get { return macroField; }
            set { macroField = value; }
        }

        [XmlAttribute]
        public string textlink
        {
            get { return textlinkField; }
            set { textlinkField = value; }
        }
        [XmlAttribute]
        public bool fLocksText
        {
            get { return fLocksTextField; }
            set { fLocksTextField = value; }
        }

        [XmlAttribute]
        public bool fPublished
        {
            get { return fPublishedField; }
            set { fPublishedField = value; }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
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
        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPr = new CT_NonVisualDrawingProps();
            return this.cNvPr;
        }
        public CT_NonVisualConnectorProperties AddNewCNvCxnSpPr()
        {
            this.cNvCxnSpPr = new CT_NonVisualConnectorProperties();
            return this.cNvCxnSpPr;
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

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_GroupShape
    {
        CT_GroupShapeProperties grpSpPrField;
        CT_GroupShapeNonVisual nvGrpSpPrField;

        public void Set(CT_GroupShape groupShape)
        {
            this.grpSpPrField = groupShape.grpSpPr;
            this.nvGrpSpPrField = groupShape.nvGrpSpPr;
        }

        public CT_GroupShapeProperties AddNewGrpSpPr()
        {
            this.grpSpPrField = new CT_GroupShapeProperties();
            return this.grpSpPrField;
        }
        public CT_GroupShapeNonVisual AddNewNvGrpSpPr()
        {
            this.nvGrpSpPrField = new CT_GroupShapeNonVisual();
            return this.nvGrpSpPrField;
        }
        public CT_Connector AddNewCxnSp()
        {
            throw new NotImplementedException();
        }
        public CT_Shape AddNewSp()
        {
            throw new NotImplementedException();
        }
        public CT_Picture AddNewPic()
        {
            throw new NotImplementedException();
        }

        public CT_GroupShapeNonVisual nvGrpSpPr
        {
            get { return nvGrpSpPrField; }
            set { nvGrpSpPrField = value; }
        }
        public CT_GroupShapeProperties grpSpPr
        {
            get { return grpSpPrField; }
            set { grpSpPrField = value; }

        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_GroupShapeNonVisual
    {
        CT_NonVisualDrawingProps cNvPrField;
        CT_NonVisualGroupDrawingShapeProps cNvGrpSpPrField;

        public CT_NonVisualGroupDrawingShapeProps AddNewCNvGrpSpPr()
        {
            this.cNvGrpSpPrField = new CT_NonVisualGroupDrawingShapeProps();
            return this.cNvGrpSpPrField;
        }
        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPrField = new CT_NonVisualDrawingProps();
            return this.cNvPrField;
        }

        public CT_NonVisualDrawingProps cNvPr
        {
            get { return cNvPrField; }
            set { cNvPrField = value; }
        }
        public CT_NonVisualGroupDrawingShapeProps cNvGrpSpPr
        {
            get { return cNvGrpSpPrField; }
            set { cNvGrpSpPrField = value; }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Marker
    {
        public int col
        {
            get;
            set;
        }
        public long colOff
        {
            get;
            set;
        }
        public int row
        {
            get;
            set;
        }
        public long rowOff
        {
            get;
            set;
        }
    }
    public enum ST_EditAs
    {
        NONE,
        twoCell,
        oneCell,
        absolute
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
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



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    [XmlRoot("wsDr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", IsNullable = true)]
    public class CT_Drawing
    {
        internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Drawing));
        public const String NAMESPACE_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public const String NAMESPACE_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        // 
        //    Saved Drawings must have the following namespaces Set:
        //    <xdr:wsDr
        //        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        //        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
        //
        ////if(isNew) xmlOptions.SetSaveSyntheticDocumentElement(new QName(CT_Drawing.type.GetName().GetNamespaceURI(), "wsDr", "xdr"));
        //Dictionary<String, String> map = new Dictionary<String, String>();
        //map[NAMESPACE_A]= "a";
        //map[ST_RelationshipId.NamespaceURI]= "r";
        ////xmlOptions.SetSaveSuggestedPrefixes(map);
        internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new[] {
            //new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"), 
            new XmlQualifiedName("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
            new XmlQualifiedName("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            new XmlQualifiedName("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")});

        private List<CT_TwoCellAnchor> twoCellAnchors = new List<CT_TwoCellAnchor>();
        private List<CT_OneCellAnchor> oneCellAnchors = new List<CT_OneCellAnchor>();
        //private List<CT_AbsoulteCellAnchor> absoluteCellAnchors = new List<CT_AbsoulteCellAnchor>();

        public CT_TwoCellAnchor AddNewTwoCellAnchor()
        {
            var anchor = new CT_TwoCellAnchor();
            twoCellAnchors.Add(anchor);
            return anchor;
        }
        public int SizeOfTwoCellAnchorArray()
        {
            return twoCellAnchors.Count;
        }
        public static CT_Drawing Parse(Stream stream)
        {
            CT_Drawing result = (CT_Drawing)serializer.Deserialize(stream);
            return result;
        }

        public void Save(Stream stream)
        {
            serializer.Serialize(stream, this, namespaces);
        }

        [XmlElement("twoCellAnchor")]
        public List<CT_TwoCellAnchor> TwoCellAnchors
        {
            get { return twoCellAnchors; }
            set { twoCellAnchors = value; }
        }

        [XmlElement("oneCellAnchor")]
        public List<CT_OneCellAnchor> OneCellAnchors
        {
            get { return oneCellAnchors; }
            set { oneCellAnchors = value; }
        }

        //[XmlElement("absoluteAnchor")]
        //public List<CT_TwoCellAnchor> AbsoluteAnchors
        //{
        //    get { return absoluteAnchors; }
        //    set { absoluteAnchors = value; }
        //}

        public void Set(CT_Drawing cT_Drawing)
        {
            throw new NotImplementedException();
        }

        public object SizeOfAbsoluteAnchorArray()
        {
            throw new NotImplementedException();
        }

        public object SizeOfOneCellAnchorArray()
        {
            throw new NotImplementedException();
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_LegacyDrawing
    {

        private string idField;

        [XmlAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }
    }
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_OneCellAnchor
    {
        private CT_Marker fromField = new CT_Marker();
        private CT_PositiveSize2D extField; //= new CT_PositiveSize2D();
        private CT_AnchorClientData clientDataField = new CT_AnchorClientData(); // 1..1 element
        [XmlElement]
        public CT_Marker from
        {
            get { return fromField; }
            set { fromField = value; }
        }

        [XmlElement]
        public CT_PositiveSize2D ext
        {
            get { return this.extField; }
            set { this.extField = value; }
        }
        public CT_AnchorClientData AddNewClientData()
        {
            this.clientDataField = new CT_AnchorClientData();
            return this.clientDataField;
        }
        [XmlElement]
        public CT_AnchorClientData clientData
        {
            get { return clientDataField; }
            set { clientDataField = value; }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_TwoCellAnchor // : EG_Anchor - was empty interface
    {
        private CT_Marker fromField = new CT_Marker(); // 1..1 element
        private CT_Marker toField = new CT_Marker(); // 1..1 element
        // 1..1 element choice - one of CT_Shape, CT_GroupShape, CT_GraphicalObjectFrame, CT_Connector or CT_Picture
        private CT_Shape shapeField = null;
        private CT_GroupShape groupShapeField = null;
        private CT_GraphicalObjectFrame graphicalObjectField = null;
        private CT_Connector connectorField = null;
        private CT_Picture pictureField = null;

        private CT_AnchorClientData clientDataField = new CT_AnchorClientData(); // 1..1 element

        private ST_EditAs editAsField = ST_EditAs.NONE; // 0..1 attribute

        public CT_Shape AddNewSp()
        {
            shapeField = new CT_Shape();
            return shapeField;
        }

        public CT_GroupShape AddNewGrpSp()
        {
            groupShapeField = new CT_GroupShape();
            return groupShapeField;
        }

        public CT_GraphicalObjectFrame AddNewGraphicFrame()
        {
            graphicalObjectField = new CT_GraphicalObjectFrame();
            return graphicalObjectField;
        }

        public CT_Connector AddNewCxnSp()
        {
            connectorField = new CT_Connector();
            return connectorField;
        }

        public CT_Picture AddNewPic()
        {
            pictureField = new CT_Picture();
            return pictureField;
        }

        public CT_AnchorClientData AddNewClientData()
        {
            this.clientDataField = new CT_AnchorClientData();
            return this.clientDataField;
        }

        [XmlAttribute]
        public ST_EditAs editAs
        {
            get { return editAsField; }
            set { editAsField = value; }
        }
        bool editAsSpecifiedField = false;
        [XmlIgnore]
        public bool editAsSpecified
        {
            get { return editAsSpecifiedField; }
            set { editAsSpecifiedField = value; }
        }

        [XmlElement]
        public CT_Marker from
        {
            get { return fromField; }
            set { fromField = value; }
        }

        [XmlElement]
        public CT_Marker to
        {
            get { return toField; }
            set { toField = value; }
        }

        #region Choice - one of CT_Shape, CT_GroupShape, CT_GraphicalObjectFrame, CT_Connector or CT_Picture

        [XmlElement]
        public CT_Shape sp
        {
            get { return shapeField; }
            set { shapeField = value; }
        }
        //bool shapeSpecifiedField = false;
        //[XmlIgnore]
        //public bool spSpecified
        //{
        //    get { return shapeSpecifiedField; }
        //    set { shapeSpecifiedField = value; }
        //}

        [XmlElement]
        public CT_GroupShape groupShape
        {
            get { return groupShapeField; }
            set { groupShapeField = value; }
        }
        //bool groupShapeSpecifiedField = false;
        //[XmlIgnore]
        //public bool groupShapeSpecified
        //{
        //    get { return groupShapeSpecifiedField; }
        //    set { groupShapeSpecifiedField = value; }
        //}

        [XmlElement]
        public CT_GraphicalObjectFrame graphicFrame
        {
            get { return graphicalObjectField; }
            set { graphicalObjectField = value; }
        }
        //bool graphicalObjectSpecifiedField = false;
        //[XmlIgnore]
        //public bool graphicalObjectSpecified
        //{
        //    get { return graphicalObjectSpecifiedField; }
        //    set { graphicalObjectSpecifiedField = value; }
        //}

        [XmlElement]
        public CT_Connector connector
        {
            get { return connectorField; }
            set { connectorField = value; }
        }
        //bool connectorSpecifiedField = false;
        //[XmlIgnore]
        //public bool connectorSpecified
        //{
        //    get { return connectorSpecifiedField; }
        //    set { connectorSpecifiedField = value; }
        //}

        [XmlElement("pic")]
        public CT_Picture picture
        {
            get { return pictureField; }
            set { pictureField = value; }
        }
        //bool pictureSpecifiedField = false;
        //[XmlIgnore]
        //public bool pictureSpecified
        //{
        //    get { return pictureSpecifiedField; }
        //    set { pictureSpecifiedField = value; }
        //}

        #endregion Choice - one of CT_Shape, CT_GroupShape, CT_GraphicalObjectFrame, CT_Connector or CT_Picture

        [XmlElement]
        public CT_AnchorClientData clientData
        {
            get { return clientDataField; }
            set { clientDataField = value; }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Connector // empty interface: EG_ObjectChoices
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
        public void Set(CT_Connector obj)
        {
            //throw new NotImplementedException();
            this.macroField = obj.macro;
            this.fPublishField = obj.fPublished;
            this.spPrField = obj.spPr;
            this.styleField = obj.style;
            this.nvCxnSpPrField = obj.nvCxnSpPr;
        }
        public CT_ConnectorNonVisual AddNewNvCxnSpPr()
        {
            this.nvCxnSpPr = new CT_ConnectorNonVisual();
            return nvCxnSpPr;
        }
        public CT_ShapeProperties AddNewSpPr()
        {
            this.spPrField = new CT_ShapeProperties();
            return spPrField;
        }
        public CT_ShapeStyle AddNewStyle()
        {
            this.styleField = new CT_ShapeStyle();
            return this.styleField;
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
        [XmlAttribute]
        public string macro
        {
            get { return this.macroField; }
            set { this.macroField = value; }
        }
        [XmlAttribute]
        public bool fPublished
        {
            get { return this.fPublishField; }
            set { this.fPublishField = value; }
        }
    }

    
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Picture // empty interface: EG_ObjectChoices
    {
        private CT_PictureNonVisual nvPicPrField = new CT_PictureNonVisual();        //  draw-ssdraw: 1..1 
        private CT_BlipFillProperties blipFillField = new CT_BlipFillProperties();   //  draw-ssdraw: 1..1 
        private CT_ShapeProperties spPrField = new CT_ShapeProperties();             //  draw-ssdraw: 1..1 
        private CT_ShapeStyle styleField = null; // 0..1

        private string macroField = null;
        private bool fPublishedField = false;

        [XmlElement]
        public CT_PictureNonVisual nvPicPr
        {
            get { return this.nvPicPrField; }
            set { this.nvPicPrField = value; }
        }

        [XmlElement]
        public CT_BlipFillProperties blipFill
        {
            get { return this.blipFillField; }
            set { this.blipFillField = value; }
        }

        [XmlElement]
        public CT_ShapeProperties spPr
        {
            get { return this.spPrField; }
            set { this.spPrField = value; }
        }

        [XmlElement]
        public CT_ShapeStyle style
        {
            get {
                return this.styleField;
            }
            set {
                this.styleField = value;
            }
        }
        private bool styleSpecifiedField = false;
        [XmlIgnore]
        public bool styleSpecified
        {
            get { return styleSpecifiedField; }
            set { styleSpecifiedField = value; }
        }
        
        [XmlAttribute]
        public string macro
        {
            get { return macroField; }
            set { macroField = value; }
        }
        private bool macroSpecifiedField = false;
        [XmlIgnore]
        public bool macroSpecified
        {
            get { return macroSpecifiedField; }
            set { macroSpecifiedField = value; }
        }

        [XmlAttribute]
        public bool fPublished
        {
            get { return fPublishedField; }
            set { fPublishedField = value; }
        }
        private bool fPublishedSpecifiedField = false;
        [XmlIgnore]
        public bool fPublishedSpecified
        {
            get { return fPublishedSpecifiedField; }
            set { fPublishedSpecifiedField = value; }
        }

        public CT_PictureNonVisual AddNewNvPicPr()
        {
            nvPicPrField = new CT_PictureNonVisual();
            return this.nvPicPrField;
        }

        public CT_BlipFillProperties AddNewBlipFill()
        {
            blipFillField = new CT_BlipFillProperties();
            return this.blipFillField;
        }

        public CT_ShapeProperties AddNewSpPr()
        {
            spPrField = new CT_ShapeProperties();
            return this.spPrField;
        }

        public void Set(CT_Picture pict)
        {
            this.nvPicPr = pict.nvPicPr;
            this.spPr = pict.spPr;
            this.macro = pict.macro;
            this.macroSpecified = this.macroSpecified;
            this.style = pict.style;
            this.styleSpecified = pict.styleSpecified;
            this.fPublished = pict.fPublished;
            this.fPublishedSpecified = pict.fPublishedSpecified;
            this.blipFill = pict.blipFill;
            

        }
    }

    // see same class in different name space in Picture.cs
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_PictureNonVisual
    {

        private CT_NonVisualDrawingProps cNvPrField = new CT_NonVisualDrawingProps(); // 1..1
        private CT_NonVisualPictureProperties cNvPicPrField = new CT_NonVisualPictureProperties(); // 1..1

        public CT_NonVisualDrawingProps AddNewCNvPr()
        {
            this.cNvPrField = new CT_NonVisualDrawingProps();
            return this.cNvPrField;
        }
        public CT_NonVisualPictureProperties AddNewCNvPicPr()
        {
            this.cNvPicPrField = new CT_NonVisualPictureProperties();
            return this.cNvPicPrField;
        }

        [XmlElement]
        public CT_NonVisualDrawingProps cNvPr
        {
            get { return this.cNvPrField; }
            set { this.cNvPrField = value; }
        }


        [XmlElement]
        public CT_NonVisualPictureProperties cNvPicPr
        {
            get { return this.cNvPicPrField; }
            set { this.cNvPicPrField = value; }
        }
    }


}
