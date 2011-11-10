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
namespace NPOI.HWPF.usermodel;

using java.util.ArrayList;
using java.util.Collection;
using java.util.Collections;
using java.util.List;

using NPOI.DDF.EscherTertiaryOptRecord;

using NPOI.DDF.DefaultEscherRecordFactory;
using NPOI.DDF.EscherBSERecord;
using NPOI.DDF.EscherBlipRecord;
using NPOI.DDF.EscherContainerRecord;
using NPOI.DDF.EscherOptRecord;
using NPOI.DDF.EscherProperties;
using NPOI.DDF.EscherRecord;
using NPOI.DDF.EscherRecordFactory;
using NPOI.DDF.EscherSimpleProperty;
using NPOI.DDF.EscherSpRecord;
using NPOI.HWPF.model.EscherRecordHolder;
using NPOI.HWPF.model.FSPA;
using NPOI.HWPF.model.FSPATable;

public class OfficeDrawingsImpl : OfficeDrawings
{
    private EscherRecordHolder _escherRecordHolder;
    private FSPATable _fspaTable;
    private byte[] _mainStream;

    public OfficeDrawingsImpl( FSPATable fspaTable,
            EscherRecordHolder escherRecordHolder, byte[] mainStream )
    {
        this._fspaTable = fspaTable;
        this._escherRecordHolder = escherRecordHolder;
        this._mainStream = mainStream;
    }

    private EscherBlipRecord GetBitmapRecord( int bitmapIndex )
    {
        List<? : EscherContainerRecord> bContainers = _escherRecordHolder
                .GetBStoreContainers();
        if ( bContainers == null || bContainers.Count != 1 )
            return null;

        EscherContainerRecord bContainer = bContainers.Get( 0 );
        List<EscherRecord> bitmapRecords = bContainer.GetChildRecords();

        if ( bitmapRecords.Count < bitmapIndex )
            return null;

        EscherRecord imageRecord = bitmapRecords.Get( bitmapIndex - 1 );

        if ( imageRecord is EscherBlipRecord )
        {
            return (EscherBlipRecord) imageRecord;
        }

        if ( imageRecord is EscherBSERecord )
        {
            EscherBSERecord bseRecord = (EscherBSERecord) imageRecord;

            EscherBlipRecord blip = bseRecord.GetBlipRecord();
            if ( blip != null )
            {
                return blip;
            }

            if ( bseRecord.GetOffSet() > 0 )
            {
                /*
                 * Blip stored in delay stream, which in a word doc, is the main
                 * stream
                 */
                EscherRecordFactory recordFactory = new DefaultEscherRecordFactory();
                EscherRecord record = recordFactory.CreateRecord( _mainStream,
                        bseRecord.GetOffSet() );

                if ( record is EscherBlipRecord )
                {
                    record.FillFields( _mainStream, bseRecord.GetOffSet(),
                            recordFactory );
                    return (EscherBlipRecord) record;
                }
            }
        }

        return null;
    }

    private EscherContainerRecord GetEscherShapeRecordContainer(
            int shapeId )
    {
        for ( EscherContainerRecord spContainer : _escherRecordHolder
                .GetSpContainers() )
        {
            EscherSpRecord escherSpRecord = spContainer
                    .GetChildById( (short) 0xF00A );
            if ( escherSpRecord != null
                    && escherSpRecord.GetShapeId() == shapeId )
                return spContainer;
        }

        return null;
    }

    private OfficeDrawing GetOfficeDrawing( FSPA fspa )
    {
        return new OfficeDrawing()
        {
            public HorizontalPositioning GetHorizontalPositioning()
            {
                int value = GetTertiaryPropertyValue(
                        EscherProperties.GROUPSHAPE__POSH, -1 );

                switch ( value )
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
                        EscherProperties.GROUPSHAPE__POSRELH, -1 );

                switch ( value )
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
                EscherContainerRecord shapeDescription = GetEscherShapeRecordContainer( GetShapeId() );
                if ( shapeDescription == null )
                    return null;

                EscherOptRecord escherOptRecord = shapeDescription
                        .GetChildById( EscherOptRecord.RECORD_ID );
                if ( escherOptRecord == null )
                    return null;

                EscherSimpleProperty escherProperty = escherOptRecord
                        .Lookup( EscherProperties.BLIP__BLIPTODISPLAY );
                if ( escherProperty == null )
                    return null;

                int bitmapIndex = escherProperty.GetPropertyValue();
                EscherBlipRecord escherBlipRecord = GetBitmapRecord( bitmapIndex );
                if ( escherBlipRecord == null )
                    return null;

                return escherBlipRecord.GetPicturedata();
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

            private int GetTertiaryPropertyValue( int propertyId,
                    int defaultValue )
            {
                EscherContainerRecord shapeDescription = GetEscherShapeRecordContainer( GetShapeId() );
                if ( shapeDescription == null )
                    return defaultValue;

                EscherTertiaryOptRecord escherTertiaryOptRecord = shapeDescription
                        .GetChildById( EscherTertiaryOptRecord.RECORD_ID );
                if ( escherTertiaryOptRecord == null )
                    return defaultValue;

                EscherSimpleProperty escherProperty = escherTertiaryOptRecord
                        .Lookup( propertyId );
                if ( escherProperty == null )
                    return defaultValue;
                int value = escherProperty.GetPropertyValue();

                return value;
            }

            public VerticalPositioning GetVerticalPositioning()
            {
                int value = GetTertiaryPropertyValue(
                        EscherProperties.GROUPSHAPE__POSV, -1 );

                switch ( value )
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
                        EscherProperties.GROUPSHAPE__POSV, -1 );

                switch ( value )
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

            @Override
            public String ToString()
            {
                return "OfficeDrawingImpl: " + fspa.ToString();
            }
        };
    }

    public OfficeDrawing GetOfficeDrawingAt( int characterPosition )
    {
        FSPA fspa = _fspaTable.GetFspaFromCp( characterPosition );
        if ( fspa == null )
            return null;

        return GetOfficeDrawing( fspa );
    }

    public Collection<OfficeDrawing> GetOfficeDrawings()
    {
        List<OfficeDrawing> result = new ArrayList<OfficeDrawing>();
        for ( FSPA fspa : _fspaTable.GetShapes() )
        {
            result.Add( GetOfficeDrawing( fspa ) );
        }
        return Collections.unmodifiableList( result );
    }
}


