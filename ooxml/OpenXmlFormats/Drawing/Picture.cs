using System;
using System.ComponentModel;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml.Picture
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/picture")]
    [XmlRoot("pic", Namespace = "http://schemas.openxmlformats.org/drawingml/2006/picture", IsNullable = false)]
    // draw-pic:pic
    public class CT_Picture
    {
        private CT_PictureNonVisual nvPicPrField = new CT_PictureNonVisual();        //  draw-pic 1..1 

        private CT_BlipFillProperties blipFillField = new CT_BlipFillProperties();   //  draw-pic: 1..1 

        private CT_ShapeProperties spPrField = new CT_ShapeProperties();             //  draw-pic: 1..1 

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

        //public void Set(CT_Picture pict)
        //{
        //    throw new NotImplementedException();
        //}

    }

    // see same class in different name space in SpeedsheetDrawing.cs
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/picture")]
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
namespace NPOI.OpenXmlFormats.Dml
{

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public class CT_NonVisualPictureProperties
    {

        private CT_PictureLocking picLocksField = null;

        private CT_OfficeArtExtensionList extLstField = null;

        private bool? preferRelativeResizeField = null;

        public CT_NonVisualPictureProperties()
        {
            this.preferRelativeResizeField = true;
        }

        public CT_PictureLocking AddNewPicLocks()
        {
            this.picLocksField = new CT_PictureLocking();
            return picLocksField;
        }

        [XmlElement]
        public CT_PictureLocking picLocks
        {
            get
            {
                return this.picLocksField;
            }
            set
            {
                this.picLocksField = value;
            }
        }


        [XmlElement]
        public CT_OfficeArtExtensionList extLst
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
        [DefaultValue(true)]
        public bool preferRelativeResize
        {
            get
            {
                return null == this.preferRelativeResizeField ? true : (bool)preferRelativeResizeField;
            }
            set
            {
                this.preferRelativeResizeField = value;
            }
        }
        [XmlIgnore]
        public bool preferRelativeResizeSpecified
        {
            get { return (null != preferRelativeResizeField); }
        }

    }

    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public class CT_NonVisualDrawingProps
    {
        private CT_Hyperlink hlinkClickField = null;

        private CT_Hyperlink hlinkHoverField = null;

        private CT_OfficeArtExtensionList extLstField = null;

        private uint idField; // 1..1

        private string nameField; // 1..1

        private string descrField = null;

        private bool? hiddenField = null;


        [XmlElement(Order = 0)]
        public CT_Hyperlink hlinkClick
        {
            get
            {
                return this.hlinkClickField;
            }
            set
            {
                this.hlinkClickField = value;
            }
        }


        [XmlElement(Order = 1)]
        public CT_Hyperlink hlinkHover
        {
            get
            {
                return this.hlinkHoverField;
            }
            set
            {
                this.hlinkHoverField = value;
            }
        }


        [XmlElement(Order = 2)]
        public CT_OfficeArtExtensionList extLst
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
        public uint id
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


        [XmlAttribute]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue("")]
        public string descr
        {
            get
            {
                return null == this.descrField ? "" : descrField;
            }
            set
            {
                this.descrField = value;
            }
        }
        [XmlIgnore]
        public bool descrSpecified
        {
            get { return (null != descrField); }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool hidden
        {
            get
            {
                return null == this.hiddenField ? false : (bool)hiddenField;
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
        }
    }


    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public class CT_PictureLocking
    {

        private CT_OfficeArtExtensionList extLstField = null;

        // TODO all attributes are optional
        private bool noGrpField;

        private bool noSelectField;

        private bool noRotField;

        private bool noChangeAspectField;

        private bool noMoveField;

        private bool noResizeField;

        private bool noEditPointsField;

        private bool noAdjustHandlesField;

        private bool noChangeArrowheadsField;

        private bool noChangeShapeTypeField;

        private bool noCropField;

        public CT_PictureLocking()
        {
            this.noGrpField = false;
            this.noSelectField = false;
            this.noRotField = false;
            this.noChangeAspectField = false;
            this.noMoveField = false;
            this.noResizeField = false;
            this.noEditPointsField = false;
            this.noAdjustHandlesField = false;
            this.noChangeArrowheadsField = false;
            this.noChangeShapeTypeField = false;
            this.noCropField = false;
        }

        [XmlElement]
        public CT_OfficeArtExtensionList extLst
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
        [DefaultValue(false)]
        public bool noGrp
        {
            get
            {
                return this.noGrpField;
            }
            set
            {
                this.noGrpField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noSelect
        {
            get
            {
                return this.noSelectField;
            }
            set
            {
                this.noSelectField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noRot
        {
            get
            {
                return this.noRotField;
            }
            set
            {
                this.noRotField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeAspect
        {
            get
            {
                return this.noChangeAspectField;
            }
            set
            {
                this.noChangeAspectField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noMove
        {
            get
            {
                return this.noMoveField;
            }
            set
            {
                this.noMoveField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noResize
        {
            get
            {
                return this.noResizeField;
            }
            set
            {
                this.noResizeField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noEditPoints
        {
            get
            {
                return this.noEditPointsField;
            }
            set
            {
                this.noEditPointsField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noAdjustHandles
        {
            get
            {
                return this.noAdjustHandlesField;
            }
            set
            {
                this.noAdjustHandlesField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeArrowheads
        {
            get
            {
                return this.noChangeArrowheadsField;
            }
            set
            {
                this.noChangeArrowheadsField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noChangeShapeType
        {
            get
            {
                return this.noChangeShapeTypeField;
            }
            set
            {
                this.noChangeShapeTypeField = value;
            }
        }


        [XmlAttribute]
        [DefaultValue(false)]
        public bool noCrop
        {
            get
            {
                return this.noCropField;
            }
            set
            {
                this.noCropField = value;
            }
        }
    }

}

