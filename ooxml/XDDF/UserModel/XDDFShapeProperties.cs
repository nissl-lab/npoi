/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.XDDF.UserModel
{
    using NPOI.Util;
    using NPOI.OpenXmlFormats.Dml.Chart;
    public class XDDFShapeProperties
    {
        private CT_ShapeProperties props;

        public XDDFShapeProperties()
            : this(new CT_ShapeProperties())
        {

        }
        public XDDFShapeProperties(CT_ShapeProperties properties)
        {
            this.props = properties;
        }
        public CT_ShapeProperties GetXmlobject()
        {
            return props;
        }

        public BlackWhiteMode? GetBlackWhiteMode()
        {
            if(props.IsSetBwMode())
            {
                return BlackWhiteModeExtensions.ValueOf(props.bwMode);
            }
            else
            {
                return null;
            }
        }

        public void SetBlackWhiteMode(BlackWhiteMode? mode)
        {
            if(mode == null)
            {
                if(props.IsSetBwMode())
                {
                    props.UnsetBwMode();
                }
            }
            else
            {
                props.bwMode = mode.Value.ToST_BlackWhiteMode();
            }
        }

        //public XDDFFillProperties GetFillProperties()
        //{
        //    if(props.IsSetGradFill())
        //    {
        //        return new XDDFGradientFillProperties(props.GradFill);
        //    }
        //    else if(props.IsSetGrpFill())
        //    {
        //        return new XDDFGroupFillProperties(props.GrpFill);
        //    }
        //    else if(props.IsSetNoFill())
        //    {
        //        return new XDDFNoFillProperties(props.NoFill);
        //    }
        //    else if(props.IsSetPattFill())
        //    {
        //        return new XDDFPatternFillProperties(props.PattFill);
        //    }
        //    else if(props.IsSetBlipFill())
        //    {
        //        return new XDDFPictureFillProperties(props.BlipFill);
        //    }
        //    else if(props.IsSetSolidFill())
        //    {
        //        return new XDDFSolidFillProperties(props.SolidFill);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetFillProperties(XDDFFillProperties properties)
        //{
        //    if(props.IsSetBlipFill())
        //    {
        //        props.unsetBlipFill();
        //    }
        //    if(props.IsSetGradFill())
        //    {
        //        props.unsetGradFill();
        //    }
        //    if(props.IsSetGrpFill())
        //    {
        //        props.unsetGrpFill();
        //    }
        //    if(props.IsSetNoFill())
        //    {
        //        props.unsetNoFill();
        //    }
        //    if(props.IsSetPattFill())
        //    {
        //        props.unsetPattFill();
        //    }
        //    if(props.IsSetSolidFill())
        //    {
        //        props.unsetSolidFill();
        //    }
        //    if(properties == null)
        //    {
        //        return;
        //    }
        //    if(properties is XDDFGradientFillProperties)
        //    {
        //        props.SetGradFill(((XDDFGradientFillProperties) properties).Xmlobject);
        //    }
        //    else if(properties is XDDFGroupFillProperties)
        //    {
        //        props.SetGrpFill(((XDDFGroupFillProperties) properties).Xmlobject);
        //    }
        //    else if(properties is XDDFNoFillProperties)
        //    {
        //        props.SetNoFill(((XDDFNoFillProperties) properties).Xmlobject);
        //    }
        //    else if(properties is XDDFPatternFillProperties)
        //    {
        //        props.SetPattFill(((XDDFPatternFillProperties) properties).Xmlobject);
        //    }
        //    else if(properties is XDDFPictureFillProperties)
        //    {
        //        props.SetBlipFill(((XDDFPictureFillProperties) properties).Xmlobject);
        //    }
        //    else if(properties is XDDFSolidFillProperties)
        //    {
        //        props.SetSolidFill(((XDDFSolidFillProperties) properties).Xmlobject);
        //    }
        //}

        //public XDDFLineProperties GetLineProperties()
        //{
        //    if(props.IsSetLn())
        //    {
        //        return new XDDFLineProperties(props.Ln);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetLineProperties(XDDFLineProperties properties)
        //{
        //    if(properties == null)
        //    {
        //        if(props.IsSetLn())
        //        {
        //            props.unsetLn();
        //        }
        //    }
        //    else
        //    {
        //        props.SetLn(properties.Xmlobject);
        //    }
        //}

        //public XDDFCustomGeometry2D GetCustomGeometry2D()
        //{
        //    if(props.IsSetCustGeom())
        //    {
        //        return new XDDFCustomGeometry2D(props.CustGeom);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetCustomGeometry2D(XDDFCustomGeometry2D geometry)
        //{
        //    if(geometry == null)
        //    {
        //        if(props.IsSetCustGeom())
        //        {
        //            props.unsetCustGeom();
        //        }
        //    }
        //    else
        //    {
        //        props.SetCustGeom(geometry.Xmlobject);
        //    }
        //}

        //public XDDFPresetGeometry2D GetPresetGeometry2D()
        //{
        //    if(props.IsSetPrstGeom())
        //    {
        //        return new XDDFPresetGeometry2D(props.PrstGeom);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetPresetGeometry2D(XDDFPresetGeometry2D geometry)
        //{
        //    if(geometry == null)
        //    {
        //        if(props.IsSetPrstGeom())
        //        {
        //            props.unsetPrstGeom();
        //        }
        //    }
        //    else
        //    {
        //        props.SetPrstGeom(geometry.Xmlobject);
        //    }
        //}

        //public XDDFEffectContainer GetEffectContainer()
        //{
        //    if(props.IsSetEffectDag())
        //    {
        //        return new XDDFEffectContainer(props.EffectDag);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetEffectContainer(XDDFEffectContainer container)
        //{
        //    if(container == null)
        //    {
        //        if(props.IsSetEffectDag())
        //        {
        //            props.unsetEffectDag();
        //        }
        //    }
        //    else
        //    {
        //        props.SetEffectDag(container.Xmlobject);
        //    }
        //}

        //public XDDFEffectList GetEffectList()
        //{
        //    if(props.IsSetEffectLst())
        //    {
        //        return new XDDFEffectList(props.EffectLst);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetEffectList(XDDFEffectList list)
        //{
        //    if(list == null)
        //    {
        //        if(props.IsSetEffectLst())
        //        {
        //            props.unsetEffectLst();
        //        }
        //    }
        //    else
        //    {
        //        props.SetEffectLst(list.Xmlobject);
        //    }
        //}

        //public XDDFExtensionList GetExtensionList()
        //{
        //    if(props.IsSetExtLst())
        //    {
        //        return new XDDFExtensionList(props.ExtLst);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetExtensionList(XDDFExtensionList list)
        //{
        //    if(list == null)
        //    {
        //        if(props.IsSetExtLst())
        //        {
        //            props.unsetExtLst();
        //        }
        //    }
        //    else
        //    {
        //        props.SetExtLst(list.Xmlobject);
        //    }
        //}

        //public XDDFScene3D GetScene3D()
        //{
        //    if(props.IsSetScene3D())
        //    {
        //        return new XDDFScene3D(props.Scene3D);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetScene3D(XDDFScene3D scene)
        //{
        //    if(scene == null)
        //    {
        //        if(props.IsSetScene3D())
        //        {
        //            props.unsetScene3D();
        //        }
        //    }
        //    else
        //    {
        //        props.SetScene3D(scene.Xmlobject);
        //    }
        //}

        //public XDDFShape3D GetShape3D()
        //{
        //    if(props.IsSetSp3D())
        //    {
        //        return new XDDFShape3D(props.Sp3D);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetShape3D(XDDFShape3D shape)
        //{
        //    if(shape == null)
        //    {
        //        if(props.IsSetSp3D())
        //        {
        //            props.unsetSp3D();
        //        }
        //    }
        //    else
        //    {
        //        props.SetSp3D(shape.Xmlobject);
        //    }
        //}

        //public XDDFTransform2D GetTransform2D()
        //{
        //    if(props.IsSetXfrm())
        //    {
        //        return new XDDFTransform2D(props.Xfrm);
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public void SetTransform2D(XDDFTransform2D transform)
        //{
        //    if(transform == null)
        //    {
        //        if(props.IsSetXfrm())
        //        {
        //            props.unsetXfrm();
        //        }
        //    }
        //    else
        //    {
        //        props.SetXfrm(transform.Xmlobject);
        //    }
        //}
    }
}


