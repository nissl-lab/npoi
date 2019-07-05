/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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
using System.Collections.Generic;
using NPOI.DDF;
using NPOI.HWPF.Model;
using System.Collections;
using System.Collections.ObjectModel;
using System;
namespace NPOI.HWPF.UserModel
{

    public class OfficeDrawingsImpl : OfficeDrawings
    {
        private EscherRecordHolder _escherRecordHolder;
        private FSPATable _fspaTable;
        private byte[] _mainStream;

        public OfficeDrawingsImpl(FSPATable fspaTable,
                EscherRecordHolder escherRecordHolder, byte[] mainStream)
        {
            this._fspaTable = fspaTable;
            this._escherRecordHolder = escherRecordHolder;
            this._mainStream = mainStream;
        }

        private EscherBlipRecord GetBitmapRecord(int bitmapIndex)
        {
            List<EscherContainerRecord> bContainers = _escherRecordHolder
                    .GetBStoreContainers();
            if (bContainers == null || bContainers.Count != 1)
                return null;

            EscherContainerRecord bContainer = bContainers[0];
            IList bitmapRecords = bContainer.ChildRecords;

            if (bitmapRecords.Count < bitmapIndex)
                return null;

            EscherRecord imageRecord = (EscherRecord)bitmapRecords[bitmapIndex - 1];

            if (imageRecord is EscherBlipRecord)
            {
                return (EscherBlipRecord)imageRecord;
            }

            if (imageRecord is EscherBSERecord)
            {
                EscherBSERecord bseRecord = (EscherBSERecord)imageRecord;

                EscherBlipRecord blip = bseRecord.BlipRecord;
                if (blip != null)
                {
                    return blip;
                }

                if (bseRecord.Offset > 0)
                {
                    /*
                     * Blip stored in delay stream, which in a word doc, is the main
                     * stream
                     */
                    IEscherRecordFactory recordFactory = new DefaultEscherRecordFactory();
                    EscherRecord record = recordFactory.CreateRecord(_mainStream,
                            bseRecord.Offset);

                    if (record is EscherBlipRecord)
                    {
                        record.FillFields(_mainStream, bseRecord.Offset,
                                recordFactory);
                        return (EscherBlipRecord)record;
                    }
                }
            }

            return null;
        }

        private EscherContainerRecord GetEscherShapeRecordContainer(
                int shapeId)
        {
            foreach (EscherContainerRecord spContainer in _escherRecordHolder
                    .GetSpContainers())
            {
                EscherSpRecord escherSpRecord = (EscherSpRecord)spContainer
                        .GetChildById(unchecked((short)0xF00A));
                if (escherSpRecord != null
                        && escherSpRecord.ShapeId == shapeId)
                    return spContainer;
            }

            return null;
        }

        private OfficeDrawing GetOfficeDrawing(FSPA fspa)
        {
            return new OfficeDrawingImpl(this,fspa);
        }

        public OfficeDrawing GetOfficeDrawingAt(int characterPosition)
        {
            FSPA fspa = _fspaTable.GetFspaFromCp(characterPosition);
            if (fspa == null)
                return null;

            return GetOfficeDrawing(fspa);
        }

        public List<OfficeDrawing> GetOfficeDrawings()
        {
            List<OfficeDrawing> result = new List<OfficeDrawing>();
            foreach (FSPA fspa in _fspaTable.GetShapes())
            {
                result.Add(GetOfficeDrawing(fspa));
            }
            return result;
        }

        internal class OfficeDrawingImpl : OfficeDrawing
        {
            FSPA fspa;
            OfficeDrawingsImpl od;
            public OfficeDrawingImpl(OfficeDrawingsImpl od,FSPA fspa)
            {
                this.fspa = fspa;
                this.od = od;
            }

            public HorizontalPositioning GetHorizontalPositioning()
            {
                int value = GetTertiaryPropertyValue(
                        EscherProperties.GROUPSHAPE__POSH, -1);

                switch (value)
                {
                    case 0:
                        return HorizontalPositioning.ABSOLUTE;
                    case 1:
                        return HorizontalPositioning.LEFT;
                    case 2:
                        return HorizontalPositioning.CENTER;
                    case 3:
                        return HorizontalPositioning.RIGHT;
                    case 4:
                        return HorizontalPositioning.INSIDE;
                    case 5:
                        return HorizontalPositioning.OUTSIDE;
                }

                return HorizontalPositioning.ABSOLUTE;
            }

            public HorizontalRelativeElement GetHorizontalRelative()
            {
                int value = GetTertiaryPropertyValue(
                        EscherProperties.GROUPSHAPE__POSRELH, -1);

                switch (value)
                {
                    case 1:
                        return HorizontalRelativeElement.MARGIN;
                    case 2:
                        return HorizontalRelativeElement.PAGE;
                    case 3:
                        return HorizontalRelativeElement.TEXT;
                    case 4:
                        return HorizontalRelativeElement.CHAR;
                }

                return HorizontalRelativeElement.TEXT;
            }

            public byte[] GetPictureData()
            {
                EscherContainerRecord shapeDescription = od.GetEscherShapeRecordContainer(GetShapeId());
                if (shapeDescription == null)
                    return null;

                EscherRecord escherOptRecord = (EscherRecord)shapeDescription
                        .GetChildById(EscherOptRecord.RECORD_ID);
                if (escherOptRecord == null)
                    return null;

                EscherSimpleProperty escherProperty = (EscherSimpleProperty)((EscherOptRecord)escherOptRecord)
                        .Lookup(EscherProperties.BLIP__BLIPTODISPLAY);
                if (escherProperty == null)
                    return null;

                int bitmapIndex = escherProperty.PropertyValue;
                EscherBlipRecord escherBlipRecord = od.GetBitmapRecord(bitmapIndex);
                if (escherBlipRecord == null)
                    return null;

                return escherBlipRecord.PictureData;
            }

            public int GetRectangleBottom()
            {
                return fspa.GetYaBottom();
            }

            public int GetRectangleLeft()
            {
                return fspa.GetXaLeft();
            }

            public int GetRectangleRight()
            {
                return fspa.GetXaRight();
            }

            public int GetRectangleTop()
            {
                return fspa.GetYaTop();
            }

            public int GetShapeId()
            {
                return fspa.GetSpid();
            }

            private int GetTertiaryPropertyValue(int propertyId,
                    int defaultValue)
            {
                EscherContainerRecord shapeDescription = od.GetEscherShapeRecordContainer(GetShapeId());
                if (shapeDescription == null)
                    return defaultValue;

                EscherRecord escherTertiaryOptRecord = (EscherRecord)shapeDescription
                        .GetChildById(EscherTertiaryOptRecord.RECORD_ID);
                if (escherTertiaryOptRecord == null)
                    return defaultValue;

                EscherSimpleProperty escherProperty = (EscherSimpleProperty)((EscherOptRecord)escherTertiaryOptRecord)
                        .Lookup(propertyId);
                if (escherProperty == null)
                    return defaultValue;
                int value = escherProperty.PropertyValue;

                return value;
            }

            public VerticalPositioning GetVerticalPositioning()
            {
                int value = GetTertiaryPropertyValue(
                        EscherProperties.GROUPSHAPE__POSV, -1);

                switch (value)
                {
                    case 0:
                        return VerticalPositioning.ABSOLUTE;
                    case 1:
                        return VerticalPositioning.TOP;
                    case 2:
                        return VerticalPositioning.CENTER;
                    case 3:
                        return VerticalPositioning.BOTTOM;
                    case 4:
                        return VerticalPositioning.INSIDE;
                    case 5:
                        return VerticalPositioning.OUTSIDE;
                }

                return VerticalPositioning.ABSOLUTE;
            }

            public VerticalRelativeElement GetVerticalRelativeElement()
            {
                int value = GetTertiaryPropertyValue(
                        EscherProperties.GROUPSHAPE__POSV, -1);

                switch (value)
                {
                    case 1:
                        return VerticalRelativeElement.MARGIN;
                    case 2:
                        return VerticalRelativeElement.PAGE;
                    case 3:
                        return VerticalRelativeElement.TEXT;
                    case 4:
                        return VerticalRelativeElement.LINE;
                }

                return VerticalRelativeElement.TEXT;
            }

            public override String ToString()
            {
                return "OfficeDrawingImpl: " + fspa.ToString();
            }
        };

    }
}


