using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXmlFormats.Dml
{
    public partial class CT_ShapeProperties
    {

        private CT_Transform2D xfrmField;

        private CT_CustomGeometry2D custGeomField;

        private CT_PresetGeometry2D prstGeomField;

        private CT_NoFillProperties noFillField;

        private CT_SolidColorFillProperties solidFillField;

        private CT_GradientFillProperties gradFillField;

        private CT_BlipFillProperties blipFillField;

        private CT_PatternFillProperties pattFillField;

        private CT_GroupFillProperties grpFillField;

        private CT_LineProperties lnField;

        private CT_EffectList effectLstField;

        private CT_EffectContainer effectDagField;

        private CT_Scene3D scene3dField;

        private CT_Shape3D sp3dField;

        private CT_OfficeArtExtensionList extLstField;

        private ST_BlackWhiteMode bwModeField;

        private bool bwModeFieldSpecified;

        public CT_ShapeProperties()
        {
            this.extLstField = new CT_OfficeArtExtensionList();
            this.sp3dField = new CT_Shape3D();
            this.scene3dField = new CT_Scene3D();
            this.effectDagField = new CT_EffectContainer();
            this.effectLstField = new CT_EffectList();
            this.lnField = new CT_LineProperties();
            this.grpFillField = new CT_GroupFillProperties();
            this.pattFillField = new CT_PatternFillProperties();
            this.blipFillField = new CT_BlipFillProperties();
            this.gradFillField = new CT_GradientFillProperties();
            this.solidFillField = new CT_SolidColorFillProperties();
            this.noFillField = new CT_NoFillProperties();
            this.prstGeomField = new CT_PresetGeometry2D();
            this.custGeomField = new CT_CustomGeometry2D();
            this.xfrmField = new CT_Transform2D();
        }

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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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
            this.xfrmField = new CT_GroupTransform2D();
        }

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

        [System.Xml.Serialization.XmlIgnoreAttribute()]
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
