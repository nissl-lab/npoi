using System;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Dml
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main", IsNullable = true)]
    public partial class CT_ShapeProperties
    {

        private CT_Transform2D xfrmField = null;

        private CT_CustomGeometry2D custGeomField = null;

        private CT_PresetGeometry2D prstGeomField = null;

        private CT_NoFillProperties noFillField = null;

        private CT_SolidColorFillProperties solidFillField = null;

        private CT_GradientFillProperties gradFillField = null;

        private CT_BlipFillProperties blipFillField = null;

        private CT_PatternFillProperties pattFillField = null;

        private CT_GroupFillProperties grpFillField = null;

        private CT_LineProperties lnField = null;

        private CT_EffectList effectLstField = null;

        private CT_EffectContainer effectDagField = null;

        private CT_Scene3D scene3dField = null;

        private CT_Shape3D sp3dField = null;

        private CT_OfficeArtExtensionList extLstField = null;

        private ST_BlackWhiteMode bwModeField = ST_BlackWhiteMode.NONE;


        //public CT_ShapeProperties()
        //{
        //    this.extLstField = new CT_OfficeArtExtensionList();
        //    this.sp3dField = new CT_Shape3D();
        //    this.scene3dField = new CT_Scene3D();
        //    this.effectDagField = new CT_EffectContainer();
        //    this.effectLstField = new CT_EffectList();
        //    //this.lnField = new CT_LineProperties();
        //    this.grpFillField = new CT_GroupFillProperties();
        //    this.pattFillField = new CT_PatternFillProperties();
        //    this.blipFillField = new CT_BlipFillProperties();
        //    this.gradFillField = new CT_GradientFillProperties();

        //    this.noFillField = new CT_NoFillProperties();
        //    this.prstGeomField = new CT_PresetGeometry2D();
        //    this.custGeomField = new CT_CustomGeometry2D();
        //    //this.xfrmField = new CT_Transform2D();
        //}
        public CT_PresetGeometry2D AddNewPrstGeom()
        {
            this.prstGeomField = new CT_PresetGeometry2D();
            return this.prstGeomField;
        }
        public CT_Transform2D AddNewXfrm()
        {
            this.xfrmField = new CT_Transform2D();
            return this.xfrmField;
        }
        public CT_SolidColorFillProperties AddNewSolidFill()
        {
            this.solidFillField = new CT_SolidColorFillProperties();
            return this.solidFillField;
        }
        public bool IsSetPattFill()
        {
            return this.pattFillField != null;
        }
        public bool IsSetSolidFill()
        {
            return this.solidFillField != null;
        }
        public bool IsSetLn()
        {
            return this.lnField != null;
        }
        public CT_LineProperties AddNewLn()
        {
            this.lnField = new CT_LineProperties();
            return lnField;
        }
        public void unsetPattFill()
        {
            this.pattFill = null;
        }
        public void unsetSolidFill()
        {
            this.solidFill = null;
        }

        [XmlElement(Order = 0)]
        public CT_Transform2D xfrm
        {
            get
            {
                return this.xfrmField;
            }
            set
            {
                this.xfrmField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_CustomGeometry2D custGeom
        {
            get
            {
                return this.custGeomField;
            }
            set
            {
                this.custGeomField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_PresetGeometry2D prstGeom
        {
            get
            {
                return this.prstGeomField;
            }
            set
            {
                this.prstGeomField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_NoFillProperties noFill
        {
            get
            {
                return this.noFillField;
            }
            set
            {
                this.noFillField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_SolidColorFillProperties solidFill
        {
            get
            {
                return this.solidFillField;
            }
            set
            {
                this.solidFillField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_GradientFillProperties gradFill
        {
            get
            {
                return this.gradFillField;
            }
            set
            {
                this.gradFillField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_BlipFillProperties blipFill
        {
            get
            {
                return this.blipFillField;
            }
            set
            {
                this.blipFillField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_PatternFillProperties pattFill
        {
            get
            {
                return this.pattFillField;
            }
            set
            {
                this.pattFillField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_GroupFillProperties grpFill
        {
            get
            {
                return this.grpFillField;
            }
            set
            {
                this.grpFillField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_LineProperties ln
        {
            get
            {
                return this.lnField;
            }
            set
            {
                this.lnField = value;
            }
        }

        [XmlElement(Order = 10)]
        public CT_EffectList effectLst
        {
            get
            {
                return this.effectLstField;
            }
            set
            {
                this.effectLstField = value;
            }
        }

        [XmlElement(Order = 11)]
        public CT_EffectContainer effectDag
        {
            get
            {
                return this.effectDagField;
            }
            set
            {
                this.effectDagField = value;
            }
        }

        [XmlElement(Order = 12)]
        public CT_Scene3D scene3d
        {
            get
            {
                return this.scene3dField;
            }
            set
            {
                this.scene3dField = value;
            }
        }

        [XmlElement(Order = 13)]
        public CT_Shape3D sp3d
        {
            get
            {
                return this.sp3dField;
            }
            set
            {
                this.sp3dField = value;
            }
        }

        [XmlElement(Order = 14)]
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
        public ST_BlackWhiteMode bwMode
        {
            get
            {
                return this.bwModeField;
            }
            set
            {
                this.bwModeField = value;
            }
        }
        [XmlIgnore]
        public bool bwModeSpecified
        {
            get { return ST_BlackWhiteMode.NONE != this.bwModeField; }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main")]
    public partial class CT_GroupShapeProperties
    {

        private CT_GroupTransform2D xfrmField;

        private CT_NoFillProperties noFillField;

        private CT_SolidColorFillProperties solidFillField;

        private CT_GradientFillProperties gradFillField;

        private CT_BlipFillProperties blipFillField;

        private CT_PatternFillProperties pattFillField;

        private CT_GroupFillProperties grpFillField;

        private CT_EffectList effectLstField;

        private CT_EffectContainer effectDagField;

        private CT_Scene3D scene3dField;

        private CT_OfficeArtExtensionList extLstField;

        private ST_BlackWhiteMode bwModeField;

        private bool bwModeFieldSpecified;

        public CT_GroupShapeProperties()
        {
            this.extLstField = new CT_OfficeArtExtensionList();
            this.scene3dField = new CT_Scene3D();
            this.effectDagField = new CT_EffectContainer();
            this.effectLstField = new CT_EffectList();
            this.grpFillField = new CT_GroupFillProperties();
            this.pattFillField = new CT_PatternFillProperties();
            this.blipFillField = new CT_BlipFillProperties();
            this.gradFillField = new CT_GradientFillProperties();
            this.solidFillField = new CT_SolidColorFillProperties();
            this.noFillField = new CT_NoFillProperties();
            //this.xfrmField = new CT_GroupTransform2D();
        }

        public CT_GroupTransform2D AddNewXfrm()
        {
            this.xfrmField = new CT_GroupTransform2D();
            return this.xfrmField;
        }
        [XmlElement(Order = 0)]
        public CT_GroupTransform2D xfrm
        {
            get
            {
                return this.xfrmField;
            }
            set
            {
                this.xfrmField = value;
            }
        }

        [XmlElement(Order = 1)]
        public CT_NoFillProperties noFill
        {
            get
            {
                return this.noFillField;
            }
            set
            {
                this.noFillField = value;
            }
        }

        [XmlElement(Order = 2)]
        public CT_SolidColorFillProperties solidFill
        {
            get
            {
                return this.solidFillField;
            }
            set
            {
                this.solidFillField = value;
            }
        }

        [XmlElement(Order = 3)]
        public CT_GradientFillProperties gradFill
        {
            get
            {
                return this.gradFillField;
            }
            set
            {
                this.gradFillField = value;
            }
        }

        [XmlElement(Order = 4)]
        public CT_BlipFillProperties blipFill
        {
            get
            {
                return this.blipFillField;
            }
            set
            {
                this.blipFillField = value;
            }
        }

        [XmlElement(Order = 5)]
        public CT_PatternFillProperties pattFill
        {
            get
            {
                return this.pattFillField;
            }
            set
            {
                this.pattFillField = value;
            }
        }

        [XmlElement(Order = 6)]
        public CT_GroupFillProperties grpFill
        {
            get
            {
                return this.grpFillField;
            }
            set
            {
                this.grpFillField = value;
            }
        }

        [XmlElement(Order = 7)]
        public CT_EffectList effectLst
        {
            get
            {
                return this.effectLstField;
            }
            set
            {
                this.effectLstField = value;
            }
        }

        [XmlElement(Order = 8)]
        public CT_EffectContainer effectDag
        {
            get
            {
                return this.effectDagField;
            }
            set
            {
                this.effectDagField = value;
            }
        }

        [XmlElement(Order = 9)]
        public CT_Scene3D scene3d
        {
            get
            {
                return this.scene3dField;
            }
            set
            {
                this.scene3dField = value;
            }
        }

        [XmlElement(Order = 10)]
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
        public ST_BlackWhiteMode bwMode
        {
            get
            {
                return this.bwModeField;
            }
            set
            {
                this.bwModeField = value;
            }
        }

        [XmlIgnore]
        public bool bwModeSpecified
        {
            get
            {
                return this.bwModeFieldSpecified;
            }
            set
            {
                this.bwModeFieldSpecified = value;
            }
        }
    }
}
