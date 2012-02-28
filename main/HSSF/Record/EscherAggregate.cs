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

namespace NPOI.HSSF.Record
{
    using System;
    using System.Text;
    using System.Collections;
    using NPOI.DDF;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Model;
    using NPOI.Util;
    using System.Collections.Generic;

    internal class SerializationListener : EscherSerializationListener
    {
        IList spEndingOffsets;
        IList shapes;

        public SerializationListener(ref IList spEndingOffsets, ref IList shapes)
        {
            this.spEndingOffsets = spEndingOffsets;
            this.shapes = shapes;
        }

        #region EscherSerializationListener Members

        void EscherSerializationListener.BeforeRecordSerialize(int Offset, short recordId, EscherRecord record)
        {
           
        }

        void EscherSerializationListener.AfterRecordSerialize(int Offset, short recordId, int size, EscherRecord record)
        {
            if (recordId == EscherClientDataRecord.RECORD_ID || recordId == EscherTextboxRecord.RECORD_ID)
            {
                spEndingOffsets.Add(Offset);
                shapes.Add(record);
            }
        }

        #endregion
    }

    /**
     * This class Is used to aggregate the MSODRAWING and OBJ record
     * combinations.  This Is necessary due to the bizare way in which
     * these records are Serialized.  What happens Is that you Get a
     * combination of MSODRAWING -> OBJ -> MSODRAWING -> OBJ records
     * but the escher records are Serialized _across_ the MSODRAWING
     * records.
     * 
     * It Gets even worse when you start looking at TXO records.
     * 
     * So what we do with this class Is aggregate lazily.  That Is
     * we don't aggregate the MSODRAWING -> OBJ records Unless we
     * need to modify them.
     *
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class EscherAggregate : AbstractEscherHolderRecord
    {
        public const short sid = 9876;
        private static POILogger log = POILogFactory.GetLogger(typeof(EscherAggregate));

        public static short ST_MIN = (short)0;
        public static short ST_NOT_PRIMATIVE = ST_MIN;
        public static short ST_RECTANGLE = (short)1;
        public static short ST_ROUNDRECTANGLE = (short)2;
        public static short ST_ELLIPSE = (short)3;
        public static short ST_DIAMOND = (short)4;
        public static short ST_ISOCELESTRIANGLE = (short)5;
        public static short ST_RIGHTTRIANGLE = (short)6;
        public static short ST_PARALLELOGRAM = (short)7;
        public static short ST_TRAPEZOID = (short)8;
        public static short ST_HEXAGON = (short)9;
        public static short ST_OCTAGON = (short)10;
        public static short ST_PLUS = (short)11;
        public static short ST_STAR = (short)12;
        public static short ST_ARROW = (short)13;
        public static short ST_THICKARROW = (short)14;
        public static short ST_HOMEPLATE = (short)15;
        public static short ST_CUBE = (short)16;
        public static short ST_BALLOON = (short)17;
        public static short ST_SEAL = (short)18;
        public static short ST_ARC = (short)19;
        public static short ST_LINE = (short)20;
        public static short ST_PLAQUE = (short)21;
        public static short ST_CAN = (short)22;
        public static short ST_DONUT = (short)23;
        public static short ST_TEXTSIMPLE = (short)24;
        public static short ST_TEXTOCTAGON = (short)25;
        public static short ST_TEXTHEXAGON = (short)26;
        public static short ST_TEXTCURVE = (short)27;
        public static short ST_TEXTWAVE = (short)28;
        public static short ST_TEXTRING = (short)29;
        public static short ST_TEXTONCURVE = (short)30;
        public static short ST_TEXTONRING = (short)31;
        public static short ST_STRAIGHTCONNECTOR1 = (short)32;
        public static short ST_BENTCONNECTOR2 = (short)33;
        public static short ST_BENTCONNECTOR3 = (short)34;
        public static short ST_BENTCONNECTOR4 = (short)35;
        public static short ST_BENTCONNECTOR5 = (short)36;
        public static short ST_CURVEDCONNECTOR2 = (short)37;
        public static short ST_CURVEDCONNECTOR3 = (short)38;
        public static short ST_CURVEDCONNECTOR4 = (short)39;
        public static short ST_CURVEDCONNECTOR5 = (short)40;
        public static short ST_CALLOUT1 = (short)41;
        public static short ST_CALLOUT2 = (short)42;
        public static short ST_CALLOUT3 = (short)43;
        public static short ST_ACCENTCALLOUT1 = (short)44;
        public static short ST_ACCENTCALLOUT2 = (short)45;
        public static short ST_ACCENTCALLOUT3 = (short)46;
        public static short ST_BORDERCALLOUT1 = (short)47;
        public static short ST_BORDERCALLOUT2 = (short)48;
        public static short ST_BORDERCALLOUT3 = (short)49;
        public static short ST_ACCENTBORDERCALLOUT1 = (short)50;
        public static short ST_ACCENTBORDERCALLOUT2 = (short)51;
        public static short ST_ACCENTBORDERCALLOUT3 = (short)52;
        public static short ST_RIBBON = (short)53;
        public static short ST_RIBBON2 = (short)54;
        public static short ST_CHEVRON = (short)55;
        public static short ST_PENTAGON = (short)56;
        public static short ST_NOSMOKING = (short)57;
        public static short ST_SEAL8 = (short)58;
        public static short ST_SEAL16 = (short)59;
        public static short ST_SEAL32 = (short)60;
        public static short ST_WEDGERECTCALLOUT = (short)61;
        public static short ST_WEDGERRECTCALLOUT = (short)62;
        public static short ST_WEDGEELLIPSECALLOUT = (short)63;
        public static short ST_WAVE = (short)64;
        public static short ST_FOLDEDCORNER = (short)65;
        public static short ST_LEFTARROW = (short)66;
        public static short ST_DOWNARROW = (short)67;
        public static short ST_UPARROW = (short)68;
        public static short ST_LEFTRIGHTARROW = (short)69;
        public static short ST_UPDOWNARROW = (short)70;
        public static short ST_IRREGULARSEAL1 = (short)71;
        public static short ST_IRREGULARSEAL2 = (short)72;
        public static short ST_LIGHTNINGBOLT = (short)73;
        public static short ST_HEART = (short)74;
        public const short ST_PICTUREFRAME = (short)75;
        public static short ST_QUADARROW = (short)76;
        public static short ST_LEFTARROWCALLOUT = (short)77;
        public static short ST_RIGHTARROWCALLOUT = (short)78;
        public static short ST_UPARROWCALLOUT = (short)79;
        public static short ST_DOWNARROWCALLOUT = (short)80;
        public static short ST_LEFTRIGHTARROWCALLOUT = (short)81;
        public static short ST_UPDOWNARROWCALLOUT = (short)82;
        public static short ST_QUADARROWCALLOUT = (short)83;
        public static short ST_BEVEL = (short)84;
        public static short ST_LEFTBRACKET = (short)85;
        public static short ST_RIGHTBRACKET = (short)86;
        public static short ST_LEFTBRACE = (short)87;
        public static short ST_RIGHTBRACE = (short)88;
        public static short ST_LEFTUPARROW = (short)89;
        public static short ST_BENTUPARROW = (short)90;
        public static short ST_BENTARROW = (short)91;
        public static short ST_SEAL24 = (short)92;
        public static short ST_STRIPEDRIGHTARROW = (short)93;
        public static short ST_NOTCHEDRIGHTARROW = (short)94;
        public static short ST_BLOCKARC = (short)95;
        public static short ST_SMILEYFACE = (short)96;
        public static short ST_VERTICALSCROLL = (short)97;
        public static short ST_HORIZONTALSCROLL = (short)98;
        public static short ST_CIRCULARARROW = (short)99;
        public static short ST_NOTCHEDCIRCULARARROW = (short)100;
        public static short ST_UTURNARROW = (short)101;
        public static short ST_CURVEDRIGHTARROW = (short)102;
        public static short ST_CURVEDLEFTARROW = (short)103;
        public static short ST_CURVEDUPARROW = (short)104;
        public static short ST_CURVEDDOWNARROW = (short)105;
        public static short ST_CLOUDCALLOUT = (short)106;
        public static short ST_ELLIPSERIBBON = (short)107;
        public static short ST_ELLIPSERIBBON2 = (short)108;
        public static short ST_FLOWCHARTProcess = (short)109;
        public static short ST_FLOWCHARTDECISION = (short)110;
        public static short ST_FLOWCHARTINPUTOUTPUT = (short)111;
        public static short ST_FLOWCHARTPREDEFINEDProcess = (short)112;
        public static short ST_FLOWCHARTINTERNALSTORAGE = (short)113;
        public static short ST_FLOWCHARTDOCUMENT = (short)114;
        public static short ST_FLOWCHARTMULTIDOCUMENT = (short)115;
        public static short ST_FLOWCHARTTERMINATOR = (short)116;
        public static short ST_FLOWCHARTPREPARATION = (short)117;
        public static short ST_FLOWCHARTMANUALINPUT = (short)118;
        public static short ST_FLOWCHARTMANUALOPERATION = (short)119;
        public static short ST_FLOWCHARTCONNECTOR = (short)120;
        public static short ST_FLOWCHARTPUNCHEDCARD = (short)121;
        public static short ST_FLOWCHARTPUNCHEDTAPE = (short)122;
        public static short ST_FLOWCHARTSUMMINGJUNCTION = (short)123;
        public static short ST_FLOWCHARTOR = (short)124;
        public static short ST_FLOWCHARTCOLLATE = (short)125;
        public static short ST_FLOWCHARTSORT = (short)126;
        public static short ST_FLOWCHARTEXTRACT = (short)127;
        public static short ST_FLOWCHARTMERGE = (short)128;
        public static short ST_FLOWCHARTOFFLINESTORAGE = (short)129;
        public static short ST_FLOWCHARTONLINESTORAGE = (short)130;
        public static short ST_FLOWCHARTMAGNETICTAPE = (short)131;
        public static short ST_FLOWCHARTMAGNETICDISK = (short)132;
        public static short ST_FLOWCHARTMAGNETICDRUM = (short)133;
        public static short ST_FLOWCHARTDISPLAY = (short)134;
        public static short ST_FLOWCHARTDELAY = (short)135;
        public static short ST_TEXTPLAINTEXT = (short)136;
        public static short ST_TEXTSTOP = (short)137;
        public static short ST_TEXTTRIANGLE = (short)138;
        public static short ST_TEXTTRIANGLEINVERTED = (short)139;
        public static short ST_TEXTCHEVRON = (short)140;
        public static short ST_TEXTCHEVRONINVERTED = (short)141;
        public static short ST_TEXTRINGINSIDE = (short)142;
        public static short ST_TEXTRINGOUTSIDE = (short)143;
        public static short ST_TEXTARCHUPCURVE = (short)144;
        public static short ST_TEXTARCHDOWNCURVE = (short)145;
        public static short ST_TEXTCIRCLECURVE = (short)146;
        public static short ST_TEXTBUTTONCURVE = (short)147;
        public static short ST_TEXTARCHUPPOUR = (short)148;
        public static short ST_TEXTARCHDOWNPOUR = (short)149;
        public static short ST_TEXTCIRCLEPOUR = (short)150;
        public static short ST_TEXTBUTTONPOUR = (short)151;
        public static short ST_TEXTCURVEUP = (short)152;
        public static short ST_TEXTCURVEDOWN = (short)153;
        public static short ST_TEXTCASCADEUP = (short)154;
        public static short ST_TEXTCASCADEDOWN = (short)155;
        public static short ST_TEXTWAVE1 = (short)156;
        public static short ST_TEXTWAVE2 = (short)157;
        public static short ST_TEXTWAVE3 = (short)158;
        public static short ST_TEXTWAVE4 = (short)159;
        public static short ST_TEXTINFLATE = (short)160;
        public static short ST_TEXTDEFLATE = (short)161;
        public static short ST_TEXTINFLATEBOTTOM = (short)162;
        public static short ST_TEXTDEFLATEBOTTOM = (short)163;
        public static short ST_TEXTINFLATETOP = (short)164;
        public static short ST_TEXTDEFLATETOP = (short)165;
        public static short ST_TEXTDEFLATEINFLATE = (short)166;
        public static short ST_TEXTDEFLATEINFLATEDEFLATE = (short)167;
        public static short ST_TEXTFADERIGHT = (short)168;
        public static short ST_TEXTFADELEFT = (short)169;
        public static short ST_TEXTFADEUP = (short)170;
        public static short ST_TEXTFADEDOWN = (short)171;
        public static short ST_TEXTSLANTUP = (short)172;
        public static short ST_TEXTSLANTDOWN = (short)173;
        public static short ST_TEXTCANUP = (short)174;
        public static short ST_TEXTCANDOWN = (short)175;
        public static short ST_FLOWCHARTALTERNATEProcess = (short)176;
        public static short ST_FLOWCHARTOFFPAGECONNECTOR = (short)177;
        public static short ST_CALLOUT90 = (short)178;
        public static short ST_ACCENTCALLOUT90 = (short)179;
        public static short ST_BORDERCALLOUT90 = (short)180;
        public static short ST_ACCENTBORDERCALLOUT90 = (short)181;
        public static short ST_LEFTRIGHTUPARROW = (short)182;
        public static short ST_SUN = (short)183;
        public static short ST_MOON = (short)184;
        public static short ST_BRACKETPAIR = (short)185;
        public static short ST_BRACEPAIR = (short)186;
        public static short ST_SEAL4 = (short)187;
        public static short ST_DOUBLEWAVE = (short)188;
        public static short ST_ACTIONBUTTONBLANK = (short)189;
        public static short ST_ACTIONBUTTONHOME = (short)190;
        public static short ST_ACTIONBUTTONHELP = (short)191;
        public static short ST_ACTIONBUTTONINFORMATION = (short)192;
        public static short ST_ACTIONBUTTONFORWARDNEXT = (short)193;
        public static short ST_ACTIONBUTTONBACKPREVIOUS = (short)194;
        public static short ST_ACTIONBUTTONEND = (short)195;
        public static short ST_ACTIONBUTTONBEGINNING = (short)196;
        public static short ST_ACTIONBUTTONRETURN = (short)197;
        public static short ST_ACTIONBUTTONDOCUMENT = (short)198;
        public static short ST_ACTIONBUTTONSOUND = (short)199;
        public static short ST_ACTIONBUTTONMOVIE = (short)200;
        public static short ST_HOSTCONTROL = (short)201;
        public const short ST_TEXTBOX = (short)202;
        public const short ST_NIL = (short)0x0FFF;

        protected HSSFPatriarch patriarch;

        /** Maps shape container objects to their OBJ records */
        private Hashtable shapeToObj = new Hashtable();
        private DrawingManager2 drawingManager;
        private short drawingGroupId;

        /**
         * list of "tail" records that need to be Serialized after all drawing Group records
         */
        private IList tailRec = new ArrayList();

        public EscherAggregate(DrawingManager2 drawingManager)
        {
            this.drawingManager = drawingManager;
        }

        /**
         * @return  Returns the current sid.
         */
        public override short Sid
        {
            get { return sid; }
        }

        /**
         * Unused since this Is an aggregate record.  Use CreateAggregate().
         *
         * @see #CreateAggregate
         */
        public System.Collections.IList Children(byte[] data, short size, int offset)
        {
            throw new InvalidOperationException("Should not reach here");
        }

        /**
         * Calculates the string representation of this record.  This Is
         * simply a dump of all the records.
         */
        public override String ToString()
        {
            String nl = Environment.NewLine;

            StringBuilder result = new StringBuilder();
            result.Append('[').Append(RecordName).Append(']' + nl);
            for (IEnumerator iterator = EscherRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord escherRecord = (EscherRecord)iterator.Current;
                result.Append(escherRecord.ToString());
            }
            result.Append("[/").Append(RecordName).Append(']' + nl);

            return result.ToString();
        }

        internal class CustomEscherRecordFactory : DefaultEscherRecordFactory
        {
            IList shapeRecords;
            public CustomEscherRecordFactory(ref IList shapeRecords)
            {
                this.shapeRecords = shapeRecords;
            }

            public override EscherRecord CreateRecord(byte[] data, int offset)
            {
                EscherRecord r = base.CreateRecord(data, offset);
                if (r.RecordId == EscherClientDataRecord.RECORD_ID || r.RecordId == EscherTextboxRecord.RECORD_ID)
                {
                    shapeRecords.Add(r);
                }
                return r;
            }
        }

        /**
         * Collapses the drawing records into an aggregate.
         */
        public static EscherAggregate CreateAggregate(IList records, int locFirstDrawingRecord, DrawingManager2 drawingManager)
        {
            // Keep track of any shape records Created so we can match them back to the object id's.
            // Textbox objects are also treated as shape objects.
            IList shapeRecords = new ArrayList();
            EscherRecordFactory recordFactory = new CustomEscherRecordFactory(ref shapeRecords);

            // Calculate the size of the buffer
            EscherAggregate agg = new EscherAggregate(drawingManager);
            int loc = locFirstDrawingRecord;
            int dataSize = 0;
            while (loc + 1 < records.Count
                    && GetSid(records, loc) == DrawingRecord.sid
                    && IsObjectRecord(records, loc + 1))
            {
                dataSize += ((DrawingRecord)records[loc]).Data.Length;
                loc += 2;
            }

            // Create one big buffer
            byte[] buffer = new byte[dataSize];
            int offset = 0;
            loc = locFirstDrawingRecord;
            while (loc + 1 < records.Count
                    && GetSid(records, loc) == DrawingRecord.sid
                    && IsObjectRecord(records, loc + 1))
            {
                DrawingRecord drawingRecord = (DrawingRecord)records[loc];
                Array.Copy(drawingRecord.Data, 0, buffer, offset, drawingRecord.Data.Length);
                offset += drawingRecord.Data.Length;
                loc += 2;
            }

            // Decode the shapes
            //        agg.escherRecords = new ArrayList();
            int pos = 0;
            while (pos < dataSize)
            {
                EscherRecord r = recordFactory.CreateRecord(buffer, pos);
                int bytesRead = r.FillFields(buffer, pos, recordFactory);
                agg.AddEscherRecord(r);
                pos += bytesRead;
            }

            // Associate the object records with the shapes
            loc = locFirstDrawingRecord;
            int shapeIndex = 0;
            agg.shapeToObj = new Hashtable();
            while (loc + 1 < records.Count
                    && GetSid(records, loc) == DrawingRecord.sid
                    && IsObjectRecord(records, loc + 1))
            {
                Record objRecord = (Record)records[loc + 1];
                agg.shapeToObj[shapeRecords[shapeIndex++]]= objRecord;
                loc += 2;
            }

            return agg;

        }

        IList spEndingOffsets;
        IList shapes;
        /**
         * Serializes this aggregate to a byte array.  Since this Is an aggregate
         * record it will effectively Serialize the aggregated records.
         *
         * @param offset    The offset into the start of the array.
         * @param data      The byte array to Serialize to.
         * @return          The number of bytes Serialized.
         */
        public override int Serialize(int offset, byte [] data)
        {
            ConvertUserModelToRecords();

            // Determine buffer size
            IList records = EscherRecords;
            int size = GetEscherRecordSize(records);
            byte[] buffer = new byte[size];


            // Serialize escher records into one big data structure and keep note of ending offsets.
            spEndingOffsets = new ArrayList();
            shapes = new ArrayList();
            int pos = 0;
            for (IEnumerator iterator = records.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord e = (EscherRecord)iterator.Current;
                pos += e.Serialize(pos, buffer, new SerializationListener(ref spEndingOffsets,ref shapes));
            }
            // todo: fix this
            shapes.Insert(0, null);
            spEndingOffsets.Insert(0, null);

            // Split escher records into Separate MSODRAWING and OBJ, TXO records.  (We don't break on
            // the first one because it's the patriach).
            pos = offset;
            for (int i = 1; i < shapes.Count; i++)
            {
                int endOffset = (int)spEndingOffsets[i] - 1;
                int startOffset;
                if (i == 1)
                    startOffset = 0;
                else
                    startOffset = (int)spEndingOffsets[i - 1];

                // Create and Write a new MSODRAWING record
                DrawingRecord drawing = new DrawingRecord();
                byte[] drawingData = new byte[endOffset - startOffset + 1];
                Array.Copy(buffer, startOffset, drawingData, 0, drawingData.Length);
                drawing.Data=drawingData;
                int temp = drawing.Serialize(pos, data);
                pos += temp;

                // Write the matching OBJ record
                Record obj = (Record)shapeToObj[shapes[i]];
                temp = obj.Serialize(pos, data);    
                pos += temp;

            }

            // Write records that need to be Serialized after all drawing Group records
            for (int i = 0; i < tailRec.Count; i++)
            {
                Record rec = (Record)tailRec[i];
                pos += rec.Serialize(pos, data);
            }

            int bytesWritten = pos - offset;
            if (bytesWritten != RecordSize)
                throw new RecordFormatException(bytesWritten + " bytes written but RecordSize reports " + RecordSize);
            return bytesWritten;
        }

        /**
         * How many bytes do the raw escher records contain.
         * @param records   List of escher records
         * @return  the number of bytes
         */
        private int GetEscherRecordSize(IList records)
        {
            int size = 0;
            for (IEnumerator iterator = records.GetEnumerator(); iterator.MoveNext(); )
                size += ((EscherRecord)iterator.Current).RecordSize;
            return size;
        }

        /**
         * The number of bytes required to Serialize this record.
         */
        public override int RecordSize
        {
            get
            {
                ConvertUserModelToRecords();
                IList records = EscherRecords;
                int rawEscherSize = GetEscherRecordSize(records);
                int drawingRecordSize = rawEscherSize + (shapeToObj.Count) * 4;
                int objRecordSize = 0;
                for (IEnumerator iterator = shapeToObj.Values.GetEnumerator(); iterator.MoveNext(); )
                {
                    Record r = (Record)iterator.Current;
                    objRecordSize += r.RecordSize;
                }
                int tailRecordSize = 0;
                for (IEnumerator iterator = tailRec.GetEnumerator(); iterator.MoveNext(); )
                {
                    Record r = (Record)iterator.Current;
                    tailRecordSize += r.RecordSize;
                }
                return drawingRecordSize + objRecordSize + tailRecordSize;
            }
        }

        /**
         * Associates an escher record to an OBJ record or a TXO record.
         */
        public Object AssoicateShapeToObjRecord(EscherRecord r, Record objRecord)
        {
            return shapeToObj[r]= objRecord;
        }

        public HSSFPatriarch Patriarch
        {
            get { return patriarch; }
            set { this.patriarch = value; }
        }

        /**
         * Converts the Records into UserModel
         *  objects on the bound HSSFPatriarch
         */
        public void ConvertRecordsToUserModel()
        {
            if (patriarch == null)
            {
                throw new InvalidOperationException("Must call SetPatriarch() first");
            }

            // The top level container ought to have
            //  the DgRecord and the container of one container
            //  per shape Group (patriach overall first)
            EscherContainerRecord topContainer =
                (EscherContainerRecord)GetEscherContainer();
            if (topContainer == null)
            {
                return;
            }
            topContainer = (EscherContainerRecord)
                topContainer.ChildContainers[0];

            IList<EscherContainerRecord> tcc = topContainer.ChildContainers;
            if (tcc.Count == 0)
            {
                throw new InvalidOperationException("No child escher containers at the point that should hold the patriach data, and one container per top level shape!");
            }

            // First up, Get the patriach position
            // This Is in the first EscherSpgrRecord, in
            //  the first container, with a EscherSRecord too
            EscherContainerRecord patriachContainer =
                (EscherContainerRecord)tcc[0];
            EscherSpgrRecord spgr = null;
            for (IEnumerator it = patriachContainer.ChildRecords.GetEnumerator(); it.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)it.Current;
                if (r is EscherSpgrRecord)
                {
                    spgr = (EscherSpgrRecord)r;
                    break;
                }
            }
            if (spgr != null)
            {
                patriarch.SetCoordinates(
                        spgr.RectX1, spgr.RectY1,
                        spgr.RectX2, spgr.RectY2
                );
            }

            // Now Process the containers for each Group
            //  and objects
            for (int i = 1; i < tcc.Count; i++)
            {
                EscherContainerRecord shapeContainer =
                    (EscherContainerRecord)tcc[i];
                //Console.Error.WriteLine("\n\n*****\n\n");
                //Console.Error.WriteLine(shapeContainer);

                // Could be a Group, or a base object
                if (shapeContainer.RecordId == EscherContainerRecord.SPGR_CONTAINER)
                {
                    if(shapeContainer.ChildRecords.Count>0)
                    {
                        // Group
                        HSSFShapeGroup group =
                            new HSSFShapeGroup(null, new HSSFClientAnchor());
                        patriarch.Children.Add(group);

                        EscherContainerRecord groupContainer =
                            (EscherContainerRecord)shapeContainer.GetChild(0);
                        ConvertRecordsToUserModel(groupContainer, group);
                    }
                }
                else if (shapeContainer.RecordId == EscherContainerRecord.SP_CONTAINER)
                {
                    EscherSpRecord spRecord = shapeContainer.GetChildById(EscherSpRecord.RECORD_ID);
                    int type = spRecord.Options >> 4;

                    switch (type)
                    {
                        case ST_TEXTBOX:
                            // TextBox
                            HSSFTextbox box =
                                new HSSFTextbox(null, new HSSFClientAnchor());
                            patriarch.Children.Add(box);

                            ConvertRecordsToUserModel(shapeContainer, box);
                            break;
                        case ST_PICTUREFRAME:
                            // Duplicated from
                            // org.apache.poi.hslf.model.Picture.getPictureIndex()
                            EscherOptRecord opt = (EscherOptRecord)GetEscherChild(shapeContainer, EscherOptRecord.RECORD_ID);
                            EscherSimpleProperty prop = (EscherSimpleProperty)opt.Lookup(EscherProperties.BLIP__BLIPTODISPLAY);
                            if (prop != null)
                            {
                                int pictureIndex = prop.PropertyValue;
                                EscherClientAnchorRecord anchorRecord = (EscherClientAnchorRecord)GetEscherChild(shapeContainer, EscherClientAnchorRecord.RECORD_ID);
                                HSSFClientAnchor anchor = new HSSFClientAnchor();
                                anchor.Col1 = anchorRecord.Col1;
                                anchor.Col2 = anchorRecord.Col2;
                                anchor.Dx1 = anchorRecord.Dx1;
                                anchor.Dx2 = anchorRecord.Dx2;
                                anchor.Dy1 = anchorRecord.Dy1;
                                anchor.Dy2 = anchorRecord.Dy2;
                                anchor.Row1 = anchorRecord.Row1;
                                anchor.Row2 = anchorRecord.Row2;
                                HSSFPicture picture = new HSSFPicture(null, anchor);
                                picture.PictureIndex = pictureIndex;
                                patriarch.AddShape(picture);

                            }
                            break;
                    }


                }
                else
                {
                    // Base level
                    ConvertRecordsToUserModel(shapeContainer, patriarch);
                }
            }

            // Now, clear any trace of what records make up
            //  the patriarch
            // Otherwise, everything will go horribly wrong
            //  when we try to Write out again....
            //    	clearEscherRecords();
            drawingManager.GetDgg().FileIdClusters=new EscherDggRecord.FileIdCluster[0];

            // TODO: Support Converting our records
            //  back into shapes
            log.Log(POILogger.WARN, "Not Processing objects into Patriarch!");
        }

        private EscherRecord GetEscherChild(EscherContainerRecord owner, int recordId)
        {
            for (IEnumerator iterator = owner.ChildRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord escherRecord = (EscherRecord)iterator.Current;
                if (escherRecord.RecordId == recordId)
                    return escherRecord;
            }
            return null;
        }


        private void ConvertRecordsToUserModel(EscherContainerRecord shapeContainer, Object model)
        {
            for (IEnumerator it = shapeContainer.ChildRecords.GetEnumerator(); it.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)it.Current;
                if (r is EscherSpgrRecord)
                {
                    // This may be overriden by a later EscherClientAnchorRecord
                    EscherSpgrRecord spgr = (EscherSpgrRecord)r;

                    if (model is HSSFShapeGroup)
                    {
                        HSSFShapeGroup g = (HSSFShapeGroup)model;
                        g.SetCoordinates(
                                spgr.RectX1, spgr.RectY1,
                                spgr.RectX2, spgr.RectY2
                        );
                    }
                    else
                    {
                        throw new InvalidOperationException("Got top level anchor but not Processing a Group");
                    }
                }
                else if (r is EscherClientAnchorRecord)
                {
                    EscherClientAnchorRecord car = (EscherClientAnchorRecord)r;

                    if (model is HSSFShape)
                    {
                        HSSFShape g = (HSSFShape)model;
                        g.Anchor.Dx1=car.Dx1;
                        g.Anchor.Dx2=car.Dx2;
                        g.Anchor.Dy1=car.Dy1;
                        g.Anchor.Dy2=car.Dy2;
                    }
                    else
                    {
                        throw new InvalidOperationException("Got top level anchor but not Processing a Group or shape");
                    }
                }
                else if (r is EscherTextboxRecord)
                {
                    EscherTextboxRecord tbr = (EscherTextboxRecord)r;

                    // Also need to Find the TextObjectRecord too
                    // TODO
                }
                else if (r is EscherSpRecord)
                {
                    // Use flags if needed
                }
                else if (r is EscherOptRecord)
                {
                    // Use properties if needed
                }
                else
                {
                    //Console.Error.WriteLine(r);
                }
            }
        }

        public void Clear()
        {
            ClearEscherRecords();
            shapeToObj.Clear();
            //        lastShapeId = 1024;
        }

        protected override String RecordName
        {
            get { return "ESCHERAGGREGATE"; }
        }

        // =============== Private methods ========================

        private static bool IsObjectRecord(IList records, int loc)
        {
            return GetSid(records, loc) == ObjRecord.sid || GetSid(records, loc) == TextObjectRecord.sid;
        }

        private void ConvertUserModelToRecords()
        {
            if (patriarch != null)
            {
                shapeToObj.Clear();
                tailRec.Clear();
                ClearEscherRecords();
                if (patriarch.Children.Count != 0)
                {
                    ConvertPatriarch(patriarch);
                    EscherContainerRecord dgContainer = (EscherContainerRecord)GetEscherRecord(0);
                    EscherContainerRecord spgrContainer = null;
                    for (int i = 0; i < dgContainer.ChildRecords.Count; i++)
                        if (dgContainer.GetChild(i).RecordId == EscherContainerRecord.SPGR_CONTAINER)
                            spgrContainer = (EscherContainerRecord)dgContainer.GetChild(i);
                    ConvertShapes(patriarch, spgrContainer, shapeToObj);

                    patriarch = null;
                }
            }
        }

        private void ConvertShapes(HSSFShapeContainer parent, EscherContainerRecord escherParent, Hashtable shapeToObj)
        {
            if (escherParent == null) throw new ArgumentException("Parent record required");

            IList shapes = parent.Children;
            for (IEnumerator iterator = shapes.GetEnumerator(); iterator.MoveNext(); )
            {
                HSSFShape shape = (HSSFShape)iterator.Current;
                if (shape is HSSFShapeGroup)
                {
                    ConvertGroup((HSSFShapeGroup)shape, escherParent, shapeToObj);
                }
                else
                {
                    AbstractShape shapeModel = AbstractShape.CreateShape(
                            shape,
                            drawingManager.AllocateShapeId(drawingGroupId));
                    shapeToObj[FindClientData(shapeModel.SpContainer)]=shapeModel.ObjRecord;
                    if (shapeModel is TextboxShape)
                    {
                        EscherRecord escherTextbox = ((TextboxShape)shapeModel).EscherTextbox;
                        shapeToObj[escherTextbox]=((TextboxShape)shapeModel).TextObjectRecord;
                        //                    escherParent.AddChildRecord(escherTextbox);

                        if (shapeModel is CommentShape)
                        {
                            CommentShape comment = (CommentShape)shapeModel;
                            tailRec.Add(comment.NoteRecord);
                        }

                    }
                    escherParent.AddChildRecord(shapeModel.SpContainer);
                }
            }
            //        drawingManager.newCluster( (short)1 );
            //        drawingManager.newCluster( (short)2 );

        }

        private void ConvertGroup(HSSFShapeGroup shape, EscherContainerRecord escherParent, Hashtable shapeToObj)
        {
            EscherContainerRecord spgrContainer = new EscherContainerRecord();
            EscherContainerRecord spContainer = new EscherContainerRecord();
            EscherSpgrRecord spgr = new EscherSpgrRecord();
            EscherSpRecord sp = new EscherSpRecord();
            EscherOptRecord opt = new EscherOptRecord();
            EscherRecord anchor;
            EscherClientDataRecord clientData = new EscherClientDataRecord();

            spgrContainer.RecordId = EscherContainerRecord.SPGR_CONTAINER;
            spgrContainer.Options = (short)0x000F;
            spContainer.RecordId = EscherContainerRecord.SP_CONTAINER;
            spContainer.Options = (short)0x000F;
            spgr.RecordId = EscherSpgrRecord.RECORD_ID;
            spgr.Options = (short)0x0001;
            spgr.RectX1 = shape.X1;
            spgr.RectY1 = shape.Y1;
            spgr.RectX2 = shape.X2;
            spgr.RectY2 = shape.Y2;
            sp.RecordId = EscherSpRecord.RECORD_ID;
            sp.Options = (short)0x0002;
            int shapeId = drawingManager.AllocateShapeId(drawingGroupId);
            sp.ShapeId = shapeId;
            if (shape.Anchor is HSSFClientAnchor)
                sp.Flags = EscherSpRecord.FLAG_GROUP | EscherSpRecord.FLAG_HAVEANCHOR;
            else
                sp.Flags = EscherSpRecord.FLAG_GROUP | EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_CHILD;
            opt.RecordId = EscherOptRecord.RECORD_ID;
            opt.Options = (short)0x0023;
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 0x00040004));
            opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.GROUPSHAPE__PRINT, 0x00080000));

            anchor = ConvertAnchor.CreateAnchor(shape.Anchor);
            //        clientAnchor.Col1( ( (HSSFClientAnchor) shape.Anchor ).Col1 );
            //        clientAnchor.Row1( (short) ( (HSSFClientAnchor) shape.Anchor ).Row1 );
            //        clientAnchor.Dx1( (short) shape.Anchor.Dx1 );
            //        clientAnchor.Dy1( (short) shape.Anchor.Dy1 );
            //        clientAnchor.Col2( ( (HSSFClientAnchor) shape.Anchor ).Col2 );
            //        clientAnchor.Row2( (short) ( (HSSFClientAnchor) shape.Anchor ).Row2 );
            //        clientAnchor.Dx2( (short) shape.Anchor.Dx2 );
            //        clientAnchor.Dy2( (short) shape.Anchor.Dy2 );
            clientData.RecordId = (EscherClientDataRecord.RECORD_ID);
            clientData.Options = ((short)0x0000);

            spgrContainer.AddChildRecord(spContainer);
            spContainer.AddChildRecord(spgr);
            spContainer.AddChildRecord(sp);
            spContainer.AddChildRecord(opt);
            spContainer.AddChildRecord(anchor);
            spContainer.AddChildRecord(clientData);

            ObjRecord obj = new ObjRecord();
            CommonObjectDataSubRecord cmo = new CommonObjectDataSubRecord();
            cmo.ObjectType = CommonObjectType.GROUP;
            cmo.ObjectId = shapeId;
            cmo.IsLocked = true;
            cmo.IsPrintable = true;
            cmo.IsAutoFill = true;
            cmo.IsAutoline = true;
            GroupMarkerSubRecord gmo = new GroupMarkerSubRecord();
            EndSubRecord end = new EndSubRecord();
            obj.AddSubRecord(cmo);
            obj.AddSubRecord(gmo);
            obj.AddSubRecord(end);
            shapeToObj[clientData] = obj;

            escherParent.AddChildRecord(spgrContainer);

            ConvertShapes(shape, spgrContainer, shapeToObj);

        }

        private EscherRecord FindClientData(EscherContainerRecord spContainer)
        {
            for (IEnumerator iterator = spContainer.ChildRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord r = (EscherRecord)iterator.Current;
                if (r.RecordId == EscherClientDataRecord.RECORD_ID)
                    return r;
            }
            throw new ArgumentException("Can not Find client data record");
        }

        private void ConvertPatriarch(HSSFPatriarch patriarch)
        {
            EscherContainerRecord dgContainer = new EscherContainerRecord();
            EscherDgRecord dg;
            EscherContainerRecord spgrContainer = new EscherContainerRecord();
            EscherContainerRecord spContainer1 = new EscherContainerRecord();
            EscherSpgrRecord spgr = new EscherSpgrRecord();
            EscherSpRecord sp1 = new EscherSpRecord();

            dgContainer.RecordId=EscherContainerRecord.DG_CONTAINER;
            dgContainer.Options=(short)0x000F;
            dg = drawingManager.CreateDgRecord();
            drawingGroupId = dg.DrawingGroupId;
            //        dg.Options( (short) ( drawingId << 4 ) );
            //        dg.NumShapes( GetNumberOfShapes( patriarch ) );
            //        dg.LastMSOSPID( 0 );  // populated after all shape id's are assigned.
            spgrContainer.RecordId=EscherContainerRecord.SPGR_CONTAINER;
            spgrContainer.Options=(short)0x000F;
            spContainer1.RecordId=EscherContainerRecord.SP_CONTAINER;
            spContainer1.Options=(short)0x000F;
            spgr.RecordId=EscherSpgrRecord.RECORD_ID;
            spgr.Options=(short)0x0001;    // version
            spgr.RectX1=patriarch.X1;
            spgr.RectY1=patriarch.Y1;
            spgr.RectX2=patriarch.X2;
            spgr.RectY2=patriarch.Y2;
            sp1.RecordId=EscherSpRecord.RECORD_ID;
            sp1.Options=(short)0x0002;
            sp1.ShapeId=drawingManager.AllocateShapeId(dg.DrawingGroupId);
            sp1.Flags=EscherSpRecord.FLAG_GROUP | EscherSpRecord.FLAG_PATRIARCH;

            dgContainer.AddChildRecord(dg);
            dgContainer.AddChildRecord(spgrContainer);
            spgrContainer.AddChildRecord(spContainer1);
            spContainer1.AddChildRecord(spgr);
            spContainer1.AddChildRecord(sp1);

            AddEscherRecord(dgContainer);
        }

        private static short GetSid(IList records, int loc)
        {
            return ((Record)records[loc]).Sid;
        }


    }
}