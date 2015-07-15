/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using NPOI.Util;
    using NPOI.DDF;
    using NPOI.HSSF.Record;

    /// <summary>
    /// @author Glen Stampoultzis  (glens at baselinksoftware.com)
    /// </summary>
    public class HSSFPolygon : HSSFSimpleShape
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(HSSFPolygon));
        public new static short OBJECT_TYPE_MICROSOFT_OFFICE_DRAWING = 0x1E;

        public HSSFPolygon(EscherContainerRecord spContainer, ObjRecord objRecord, TextObjectRecord _textObjectRecord)
            : base(spContainer, objRecord, _textObjectRecord)
        {

        }

        public HSSFPolygon(EscherContainerRecord spContainer, ObjRecord objRecord)
            : base(spContainer, objRecord)
        {

        }

        public HSSFPolygon(HSSFShape parent, HSSFAnchor anchor)
            : base(parent, anchor)
        {

        }


        protected override TextObjectRecord CreateTextObjRecord()
        {
            return null;
        }

        /**
         * Generates the shape records for this shape.
         */
        protected override EscherContainerRecord CreateSpContainer()
        {
            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherClientDataRecord clientData = new EscherClientDataRecord();

            spContainer.RecordId = (EscherContainerRecord.SP_CONTAINER);
            spContainer.Options = ((short)0x000F);
            sp.RecordId = (EscherSpRecord.RECORD_ID);
            sp.Options = ((short)((EscherAggregate.ST_NOT_PRIMATIVE << 4) | 0x2));
            if (Parent == null)
            {
                sp.Flags = (EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE);
            }
            else
            {
                sp.Flags = (EscherSpRecord.FLAG_CHILD | EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_HASSHAPETYPE);
            }
            opt.RecordId = (EscherOptRecord.RECORD_ID);
            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.TRANSFORM__ROTATION, false, false, 0));
            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__RIGHT, false, false, 100));
            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__BOTTOM, false, false, 100));
            opt.SetEscherProperty(new EscherShapePathProperty(EscherProperties.GEOMETRY__SHAPEPATH, EscherShapePathProperty.COMPLEX));

            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.GEOMETRY__FILLOK, false, false, 0x00010001));
            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINESTARTARROWHEAD, false, false, 0x0));
            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEENDARROWHEAD, false, false, 0x0));
            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEENDCAPSTYLE, false, false, 0x0));

            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEDASHING, LINESTYLE_SOLID));
            opt.SetEscherProperty(new EscherBoolProperty(EscherProperties.LINESTYLE__NOLINEDRAWDASH, 0x00080008));
            opt.SetEscherProperty(new EscherSimpleProperty(EscherProperties.LINESTYLE__LINEWIDTH, LINEWIDTH_DEFAULT));
            opt.SetEscherProperty(new EscherRGBProperty(EscherProperties.FILL__FILLCOLOR, FILL__FILLCOLOR_DEFAULT));
            opt.SetEscherProperty(new EscherRGBProperty(EscherProperties.LINESTYLE__COLOR, LINESTYLE__COLOR_DEFAULT));
            opt.SetEscherProperty(new EscherBoolProperty(EscherProperties.FILL__NOFILLHITTEST, 1));

            opt.SetEscherProperty(new EscherBoolProperty(EscherProperties.GROUPSHAPE__PRINT, 0x080000));

            EscherRecord anchor = Anchor.GetEscherAnchor();
            clientData.RecordId = (EscherClientDataRecord.RECORD_ID);
            clientData.Options = ((short)0x0000);

            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);

            return spContainer;
        }

        /**
         * Creates the low level OBJ record for this shape.
         */
        protected override ObjRecord CreateObjRecord()
        {
            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord c = new CommonObjectDataSubRecord();
            c.ObjectType = CommonObjectType.MicrosoftOfficeDrawing;
            c.IsLocked = (true);
            c.IsPrintable = (true);
            c.IsAutoFill = (true);
            c.IsAutoline = (true);
            EndSubRecord e = new EndSubRecord();
            obj.AddSubRecord(c);
            obj.AddSubRecord(e);
            return obj;
        }

        internal override void AfterRemove(HSSFPatriarch patriarch)
        {
            patriarch.GetBoundAggregate().RemoveShapeToObjRecord(GetEscherContainer().GetChildById(EscherClientDataRecord.RECORD_ID));
        }

        /**
         * @return array of x coordinates
         */
        public int[] XPoints
        {
            get
            {
                EscherArrayProperty verticesProp = (EscherArrayProperty)GetOptRecord().Lookup(EscherProperties.GEOMETRY__VERTICES);
                if (null == verticesProp)
                {
                    return new int[] { };
                }
                int[] array = new int[verticesProp.NumberOfElementsInArray - 1];
                for (int i = 0; i < verticesProp.NumberOfElementsInArray - 1; i++)
                {
                    byte[] property = verticesProp.GetElement(i);
                    short x = LittleEndian.GetShort(property, 0);
                    array[i] = x;
                }
                return array;
            }
        }

        /**
         * @return array of y coordinates
         */
        public int[] YPoints
        {
            get
            {
                EscherArrayProperty verticesProp = (EscherArrayProperty)GetOptRecord().Lookup(EscherProperties.GEOMETRY__VERTICES);
                if (null == verticesProp)
                {
                    return new int[] { };
                }
                int[] array = new int[verticesProp.NumberOfElementsInArray - 1];
                for (int i = 0; i < verticesProp.NumberOfElementsInArray - 1; i++)
                {
                    byte[] property = verticesProp.GetElement(i);
                    short x = LittleEndian.GetShort(property, 2);
                    array[i] = x;
                }
                return array;
            }
        }

        /**
         * @param xPoints - array of x coordinates
         * @param yPoints - array of y coordinates
         */
        public void SetPoints(int[] xPoints, int[] yPoints)
        {
            if (xPoints.Length != yPoints.Length)
            {
                logger.Log(POILogger.ERROR, "xPoint.Length must be equal to yPoints.Length");
                return;
            }
            if (xPoints.Length == 0)
            {
                logger.Log(POILogger.ERROR, "HSSFPolygon must have at least one point");
            }
            EscherArrayProperty verticesProp = new EscherArrayProperty(EscherProperties.GEOMETRY__VERTICES, false, new byte[0]);
            verticesProp.NumberOfElementsInArray = (xPoints.Length + 1);
            verticesProp.NumberOfElementsInMemory = (xPoints.Length + 1);
            verticesProp.SizeOfElements = unchecked((short)(0xFFF0));
            byte[] data;
            for (int i = 0; i < xPoints.Length; i++)
            {
                data = new byte[4];
                LittleEndian.PutShort(data, 0, (short)xPoints[i]);
                LittleEndian.PutShort(data, 2, (short)yPoints[i]);
                verticesProp.SetElement(i, data);
            }
            int point = xPoints.Length;
            data = new byte[4];
            LittleEndian.PutShort(data, 0, (short)xPoints[0]);
            LittleEndian.PutShort(data, 2, (short)yPoints[0]);
            verticesProp.SetElement(point, data);
            SetPropertyValue(verticesProp);

            EscherArrayProperty segmentsProp = new EscherArrayProperty(EscherProperties.GEOMETRY__SEGMENTINFO, false, null);
            segmentsProp.SizeOfElements = (0x0002);
            segmentsProp.NumberOfElementsInArray = (xPoints.Length * 2 + 4);
            segmentsProp.NumberOfElementsInMemory = (xPoints.Length * 2 + 4);
            segmentsProp.SetElement(0, new byte[] { (byte)0x00, (byte)0x40 });
            segmentsProp.SetElement(1, new byte[] { (byte)0x00, (byte)0xAC });
            for (int i = 0; i < xPoints.Length; i++)
            {
                segmentsProp.SetElement(2 + i * 2, new byte[] { (byte)0x01, (byte)0x00 });
                segmentsProp.SetElement(3 + i * 2, new byte[] { (byte)0x00, (byte)0xAC });
            }
            segmentsProp.SetElement(segmentsProp.NumberOfElementsInArray - 2, new byte[] { (byte)0x01, (byte)0x60 });
            segmentsProp.SetElement(segmentsProp.NumberOfElementsInArray - 1, new byte[] { (byte)0x00, (byte)0x80 });
            SetPropertyValue(segmentsProp);
        }

        /**
         * Defines the width and height of the points in the polygon
         * @param width
         * @param height
         */
        public void SetPolygonDrawArea(int width, int height)
        {
            SetPropertyValue(new EscherSimpleProperty(EscherProperties.GEOMETRY__RIGHT, width));
            SetPropertyValue(new EscherSimpleProperty(EscherProperties.GEOMETRY__BOTTOM, height));
        }

        /**
         * @return shape width
         */
        public int DrawAreaWidth
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.GEOMETRY__RIGHT);
                return property == null ? 100 : property.PropertyValue;
            }
        }

        /**
         * @return shape height
         */
        public int DrawAreaHeight
        {
            get
            {
                EscherSimpleProperty property = (EscherSimpleProperty)GetOptRecord().Lookup(EscherProperties.GEOMETRY__BOTTOM);
                return property == null ? 100 : property.PropertyValue;
            }
        }
    }
}