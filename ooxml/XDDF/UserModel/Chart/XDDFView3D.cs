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

namespace NPOI.XDDF.UserModel.Chart;

using NPOI.OpenXmlFormats.Dml.Chart;

public class XDDFView3D
{
    private readonly CT_View3D view3D;

    public XDDFView3D(CT_View3D view3D)
    {
        this.view3D = view3D;
    }

    public int GetXRotationAngle()
    {
        return view3D.rotX?.val ?? 0;
    }

    public void SetXRotationAngle(int val)
    {
        if (view3D.rotX != null)
        {
            view3D.rotX.val = (sbyte)val;
        }
        else
        {
            view3D.rotX = new CT_RotX { val = (sbyte)val };
        }
    }

    public int GetYRotationAngle()
    {
        return view3D.rotY?.val ?? 0;
    }

    public void SetYRotationAngle(int val)
    {
        if (view3D.rotY != null)
        {
            view3D.rotY.val = (ushort)val;
        }
        else
        {
            view3D.rotY = new CT_RotY { val = (ushort)val };
        }
    }

    public bool GetRightAngleAxes()
    {
        return view3D.rAngAx?.val == 1;
    }

    public void SetRightAngleAxes(bool val)
    {
        if (view3D.rAngAx != null)
        {
            view3D.rAngAx.val = val ? 1 : 0;
        }
        else
        {
            view3D.rAngAx = new CT_Boolean { val = val ? 1 : 0 };
        }
    }

    public byte GetPerspectiveAngle()
    {
        return view3D.perspective?.val ?? 30;
    }

    public void SetPerspectiveAngle(byte val)
    {
        if (view3D.perspective != null)
        {
            view3D.perspective.val = val;
        }
        else
        {
            view3D.perspective = new CT_Perspective { val = val };
        }
    }

    public int GetDepthPercentVal()
    {
        return view3D.depthPercent?.val ?? 100;
    }

    public void SetDepthPercent(int val)
    {
        if (view3D.depthPercent != null)
        {
            view3D.depthPercent.val = (ushort)val;
        }
        else
        {
            view3D.depthPercent = new CT_DepthPercent { val = (ushort)val };
        }
    }

    public int GetHeightPercent()
    {
        return view3D.hPercent?.val ?? 100;
    }

    public void SetHeightPercent(int val)
    {
        if (view3D.hPercent != null)
        {
            view3D.hPercent.val = (ushort)val;
        }
        else
        {
            view3D.hPercent = new CT_HPercent { val = (ushort)val };
        }
    }
}
