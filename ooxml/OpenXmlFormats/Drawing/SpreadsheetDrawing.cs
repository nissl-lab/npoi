using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Dml.Spreadsheet
{


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

        internal void Write(StreamWriter sw)
        {
            throw new NotImplementedException();
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
        public static CT_ConnectorNonVisual Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ConnectorNonVisual ctObj = new CT_ConnectorNonVisual();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "cNvCxnSpPr")
                    ctObj.cNvCxnSpPr = CT_NonVisualConnectorProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "cNvPr")
                    ctObj.cNvPr = CT_NonVisualDrawingProps.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<xdr:{0}", nodeName));
            sw.Write(">");
            if (this.cNvCxnSpPr != null)
                this.cNvCxnSpPr.Write(sw, "cNvCxnSpPr");
            if (this.cNvPr != null)
                this.cNvPr.Write(sw, "cNvPr");
            sw.Write(string.Format("</xdr:{0}>", nodeName));
        }


    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Marker
    {
        private int _col;
        private long _colOff;
        private int _row;
        private long _rowOff;

        public int col
        {
            get 
            {
                return _col;
            }
            set 
            {
                _col = value;
            }
        }
        public long colOff
        {
            get 
            {
                return _colOff;
            }
            set
            {
                _colOff = value;
            }
        }
        public int row
        {
            get { return _row; }
            set { _row = value; }
        }
        public long rowOff
        {
            get
            {
                return _rowOff;
            }
            set
            {
                _rowOff = value;
            }
        }

        public static CT_Marker Parse(XmlNode node, XmlNamespaceManager nameSpaceManager)
        {
            CT_Marker ctMarker = new CT_Marker();
            foreach (XmlNode subnode in node.ChildNodes)
            {
                if (subnode.LocalName == "col")
                {
                    ctMarker.col = Int32.Parse(subnode.InnerText);
                }
                else if (subnode.LocalName == "colOff")
                {
                    ctMarker.colOff = Int64.Parse(subnode.InnerText);
                }
                else if (subnode.LocalName == "row")
                {
                    ctMarker.row = Int32.Parse(subnode.InnerText);
                }
                else if (subnode.LocalName == "rowOff")
                {
                    ctMarker.rowOff = Int64.Parse(subnode.InnerText);
                }
            }
            return ctMarker;
        }

        public override string ToString()
        {
            StringBuilder sb=new StringBuilder();
            using(StringWriter sw =new StringWriter(sb))
            {
                sw.Write("<xdr:col>");
                sw.Write(this.col.ToString());
                sw.Write("</xdr:col>");
                sw.Write("<xdr:colOff>");
                sw.Write(this.colOff.ToString());
                sw.Write("</xdr:colOff>");
                sw.Write("<xdr:row>");
                sw.Write(this.row.ToString());
                sw.Write("</xdr:row>");
                sw.Write("<xdr:rowOff>");
                sw.Write(this.rowOff.ToString());
                sw.Write("</xdr:rowOff>");
            }
            return sb.ToString();
        }

        internal void Write(StreamWriter sw)
        {
            sw.Write("<xdr:col>");
            sw.Write(this.col.ToString());
            sw.Write("</xdr:col>");
            sw.Write("<xdr:colOff>");
            sw.Write(this.colOff.ToString());
            sw.Write("</xdr:colOff>");
            sw.Write("<xdr:row>");
            sw.Write(this.row.ToString());
            sw.Write("</xdr:row>");
            sw.Write("<xdr:rowOff>");
            sw.Write(this.rowOff.ToString());
            sw.Write("</xdr:rowOff>");
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
        bool _fLocksWithSheet;
        bool _fPrintsWithSheet;
        [XmlAttribute]
        public bool fLocksWithSheet
        {
            get
            {
                return _fLocksWithSheet;
            }
            set
            {
                _fLocksWithSheet = value;
            }
        }
        [XmlAttribute]
        public bool fPrintsWithSheet
        {
            get
            {
                return _fPrintsWithSheet;
            }
            set
            {
                _fPrintsWithSheet = value;
            }
        }
    }



    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    [XmlRoot("wsDr", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", IsNullable = true)]
    public class CT_Drawing
    {
        private List<IEG_Anchor> cellAnchors = new List<IEG_Anchor>();
        //private List<CT_AbsoulteCellAnchor> absoluteCellAnchors = new List<CT_AbsoulteCellAnchor>();

        public CT_TwoCellAnchor AddNewTwoCellAnchor()
        {
            CT_TwoCellAnchor anchor = new CT_TwoCellAnchor();
            cellAnchors.Add(anchor);
            return anchor;
        }
        public int SizeOfTwoCellAnchorArray()
        {
            int count = 0;
            foreach (IEG_Anchor anchor in cellAnchors)
            {
                if (anchor is CT_TwoCellAnchor)
                {
                    count++;
                }
            }
            return count;
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sw.Write("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
                foreach (IEG_Anchor anchor in this.cellAnchors)
                {
                    anchor.Write(sw);
                }
                sw.Write("</xdr:wsDr>");
            }
        }

        [XmlIgnore]
        public List<IEG_Anchor> CellAnchors
        {
            get { return cellAnchors; }
            set { cellAnchors = value; }
        }

        //[XmlElement("absoluteAnchor")]
        //public List<CT_TwoCellAnchor> AbsoluteAnchors
        //{
        //    get { return absoluteAnchors; }
        //    set { absoluteAnchors = value; }
        //}

        public void Set(CT_Drawing ctDrawing)
        {
            this.cellAnchors.Clear();
            foreach (IEG_Anchor anchor in ctDrawing.cellAnchors)
            {
                this.cellAnchors.Add(anchor);
            }
        }

        public int SizeOfAbsoluteAnchorArray()
        {
            return 0;
        }

        public int SizeOfOneCellAnchorArray()
        {
            int count = 0;
            foreach (IEG_Anchor anchor in cellAnchors)
            {
                if (anchor is CT_OneCellAnchor)
                {
                    count++;
                }
            }
            return count;
        }

        public static CT_Drawing Parse(XmlDocument xmldoc, XmlNamespaceManager namespaceManager)
        {
            XmlNodeList cellanchorNodes = xmldoc.SelectNodes("//*",namespaceManager);
            CT_Drawing ctDrawing = new CT_Drawing();
            foreach (XmlNode node in cellanchorNodes)
            {
                if (node.LocalName == "twoCellAnchor")
                {
                    CT_TwoCellAnchor twoCellAnchor = CT_TwoCellAnchor.Parse(node, namespaceManager);
                    ctDrawing.cellAnchors.Add(twoCellAnchor);
                }
                else if (node.LocalName == "oneCellAnchor")
                {
                    CT_OneCellAnchor oneCellAnchor = CT_OneCellAnchor.Parse(node, namespaceManager);
                    ctDrawing.cellAnchors.Add(oneCellAnchor);
                }
            }
            return ctDrawing;
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
    public class CT_OneCellAnchor: IEG_Anchor
    {
        private CT_Marker fromField = new CT_Marker();
        private CT_PositiveSize2D extField; //= new CT_PositiveSize2D();
        private CT_AnchorClientData clientDataField = new CT_AnchorClientData(); // 1..1 element
        private CT_Shape shapeField = null;
        private CT_GroupShape groupShapeField = null;
        private CT_GraphicalObjectFrame graphicalObjectField = null;
        private CT_Connector connectorField = null;
        private CT_Picture pictureField = null;
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
        public CT_Shape sp
        {
            get { return shapeField; }
            set { shapeField = value; }
        }
        public CT_GroupShape groupShape
        {
            get { return groupShapeField; }
            set { groupShapeField = value; }
        }
        public CT_GraphicalObjectFrame graphicFrame
        {
            get { return graphicalObjectField; }
            set { graphicalObjectField = value; }
        }
        public CT_Connector connector
        {
            get { return connectorField; }
            set { connectorField = value; }
        }
        public CT_Picture picture
        {
            get { return pictureField; }
            set { pictureField = value; }
        }

        public void Write(StreamWriter sw)
        {
            sw.Write("<xdr:oneCellAnchor>");
            sw.Write("<xdr:from>");
            this.from.Write(sw);
            sw.Write("</xdr:from>");
            if (this.sp != null)
                sp.Write(sw);
            sw.Write("</xdr:oneCellAnchor>");
        }

        internal static CT_OneCellAnchor Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_OneCellAnchor oneCellAnchor = new CT_OneCellAnchor();
            oneCellAnchor.from = CT_Marker.Parse(node.FirstChild, namespaceManager);
            CT_Shape ctShape = CT_Shape.Parse(node.SelectSingleNode("xdr:sp", namespaceManager), namespaceManager);
            if (ctShape != null)
            {
                oneCellAnchor.sp = ctShape;
            }
            return oneCellAnchor;
        }
    }

    public interface IEG_Anchor
    {
        CT_Shape sp { get; set; }
        CT_Connector connector { get; set; }
        CT_GraphicalObjectFrame graphicFrame { get; set; }
        CT_Picture picture { get; set; }
        CT_GroupShape groupShape { get; set; }
        CT_AnchorClientData clientData { get; set; }
        void Write(StreamWriter sw);
    }
    public class CT_AbsoluteCellAnchor : IEG_Anchor
    {
        CT_Point2D posField;
        CT_PositiveSize2D extField;
        CT_AnchorClientData clientDataField;
        private CT_Shape shapeField = null;
        private CT_GroupShape groupShapeField = null;
        private CT_GraphicalObjectFrame graphicalObjectField = null;
        private CT_Connector connectorField = null;
        private CT_Picture pictureField = null;

        public CT_AnchorClientData clientData
        {
            get { return clientDataField; }
            set { clientDataField = value; }
        }

        public CT_Point2D AddNewOff()
        {
            this.posField = new CT_Point2D();
            return this.posField;
        }

        public CT_Point2D pos
        {
            get
            {
                return this.posField;
            }
            set
            {
                this.posField = value;
            }
        }
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
        public CT_Shape sp
        {
            get { return shapeField; }
            set { shapeField = value; }
        }
        public CT_GroupShape groupShape
        {
            get { return groupShapeField; }
            set { groupShapeField = value; }
        }
        public CT_GraphicalObjectFrame graphicFrame
        {
            get { return graphicalObjectField; }
            set { graphicalObjectField = value; }
        }
        public CT_Connector connector
        {
            get { return connectorField; }
            set { connectorField = value; }
        }
        public CT_Picture picture
        {
            get { return pictureField; }
            set { pictureField = value; }
        }
        public void Write(StreamWriter sw)
        { 
            
        }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_TwoCellAnchor : IEG_Anchor //- was empty interface
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
        [XmlElement]
        public CT_GroupShape groupShape
        {
            get { return groupShapeField; }
            set { groupShapeField = value; }
        }

        [XmlElement]
        public CT_GraphicalObjectFrame graphicFrame
        {
            get { return graphicalObjectField; }
            set { graphicalObjectField = value; }
        }

        [XmlElement]
        public CT_Connector connector
        {
            get { return connectorField; }
            set { connectorField = value; }
        }

        [XmlElement("pic")]
        public CT_Picture picture
        {
            get { return pictureField; }
            set { pictureField = value; }
        }

        #endregion Choice - one of CT_Shape, CT_GroupShape, CT_GraphicalObjectFrame, CT_Connector or CT_Picture

        [XmlElement]
        public CT_AnchorClientData clientData
        {
            get { return clientDataField; }
            set { clientDataField = value; }
        }


        public void Write(StreamWriter sw)
        {
            sw.Write("<xdr:twoCellAnchor");
            if(this.editAsField!= ST_EditAs.NONE)
                sw.Write(string.Format(" editAs=\"{0}\"",this.editAsField.ToString()));
            sw.Write(">");
            sw.Write("<xdr:from>");
            this.from.Write(sw);
            sw.Write("</xdr:from>");
            sw.Write("<xdr:to>");
            this.to.Write(sw);
            sw.Write("</xdr:to>");
            if (this.picture != null)
                picture.Write(sw);
            if (this.sp != null)
                sp.Write(sw);
            //else if (this.connector != null)
            //    this.connector.Write(sw, "cxnSp");
            else if (this.groupShape != null)
                this.groupShape.Write(sw);
            else if (this.graphicalObjectField != null)
                this.graphicalObjectField.Write(sw);
            sw.Write("</xdr:twoCellAnchor>");
        }

        internal static CT_TwoCellAnchor Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_TwoCellAnchor twoCellAnchor = new CT_TwoCellAnchor();
            twoCellAnchor.from = CT_Marker.Parse(node.FirstChild, namespaceManager);
            twoCellAnchor.to = CT_Marker.Parse(node.FirstChild.NextSibling, namespaceManager);
            if (node.Attributes["editAs"] != null)
                twoCellAnchor.editAs = (ST_EditAs)Enum.Parse(typeof(ST_EditAs), node.Attributes["editAs"].Value);

            CT_Shape ctShape = CT_Shape.Parse(node.SelectSingleNode("xdr:sp", namespaceManager), namespaceManager);
            twoCellAnchor.sp = ctShape;
            CT_Picture ctPic = CT_Picture.Parse(node.SelectSingleNode("xdr:pic", namespaceManager), namespaceManager);
            twoCellAnchor.picture = ctPic;
            CT_Connector ctConnector = CT_Connector.Parse(node.SelectSingleNode("xdr:cxnSp", namespaceManager), namespaceManager);
            twoCellAnchor.connector = ctConnector;
            //TODO::parse groupshape
            CT_GroupShape ctGroupShape = CT_GroupShape.Parse(node.SelectSingleNode("xdr:grpSp", namespaceManager), namespaceManager);
            twoCellAnchor.groupShape = ctGroupShape;
            return twoCellAnchor;
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")]
    public class CT_Connector // empty interface: EG_ObjectChoices
    {
        string macroField;
        bool fPublishedField;
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
            this.macroField = obj.macro;
            this.fPublishedField = obj.fPublished;
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
            get { return this.fPublishedField; }
            set { this.fPublishedField = value; }
        }

        internal static CT_Connector Parse(XmlNode xmlNode, XmlNamespaceManager namespaceManager)
        {
            throw new NotImplementedException();
        }
    }

    

}
