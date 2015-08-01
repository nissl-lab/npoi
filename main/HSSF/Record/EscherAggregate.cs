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
    using System.IO;

    internal class SerializationListener : EscherSerializationListener
    {
        IList<int> spEndingOffsets;
        IList<EscherRecord> records;
        EscherRecord record;

        public SerializationListener(IList<int> spEndingOffsets, IList<EscherRecord> records, EscherRecord e)
        {
            this.spEndingOffsets = spEndingOffsets;
            this.records = records;
            this.record = e;
        }

        #region EscherSerializationListener Members

        void EscherSerializationListener.BeforeRecordSerialize(int offset, short recordId, EscherRecord record)
        {
           
        }

        void EscherSerializationListener.AfterRecordSerialize(int offset, short recordId, int size, EscherRecord record)
        {
            if (recordId == EscherClientDataRecord.RECORD_ID || recordId == EscherTextboxRecord.RECORD_ID)
            {
                spEndingOffsets.Add(offset);
                records.Add(record);
            }
        }

        #endregion
    }
    internal class RecordSizeListener : EscherSerializationListener
    {
        IList<int> spEndingOffsets;
        EscherRecord record;

        public RecordSizeListener(IList<int> spEndingOffsets, EscherRecord e)
        {
            this.spEndingOffsets = spEndingOffsets;
            this.record = e;
        }

        #region EscherSerializationListener Members

        void EscherSerializationListener.BeforeRecordSerialize(int offset, short recordId, EscherRecord record)
        {

        }

        void EscherSerializationListener.AfterRecordSerialize(int offset, short recordId, int size, EscherRecord record)
        {
            if (recordId == EscherClientDataRecord.RECORD_ID || recordId == EscherTextboxRecord.RECORD_ID)
            {
                spEndingOffsets.Add(offset);
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
     * At first document contains 4 types of records which belong to drawing layer.
     * There are can be such sequence of record:
     * <p/>
     * DrawingRecord
     * ContinueRecord
     * ...
     * ContinueRecord
     * ObjRecord | TextObjectRecord
     * .....
     * ContinueRecord
     * ...
     * ContinueRecord
     * ObjRecord | TextObjectRecord
     * NoteRecord
     * ...
     * NoteRecord
     * <p/>
     * To work with shapes we have to read data from Drawing and Continue records into single array of bytes and
     * build escher(office art) records tree from this array.
     * Each shape in drawing layer matches corresponding ObjRecord
     * Each textbox matches corresponding TextObjectRecord
     * <p/>
     * ObjRecord contains information about shape. Thus each ObjRecord corresponds EscherContainerRecord(SPGR)
     * <p/>
     * EscherAggrefate contains also NoteRecords
     * NoteRecords must be serial
     * 
     * @author Glen Stampoultzis (glens at apache.org)
     */
    public class EscherAggregate : AbstractEscherHolderRecord
    {
        public const short sid = 9876;
        private static POILogger log = POILogFactory.GetLogger(typeof(EscherAggregate));

        public const short ST_MIN = (short)0;
        public const short ST_NOT_PRIMATIVE = ST_MIN;
        public const short ST_RECTANGLE = (short)1;
        public const short ST_ROUNDRECTANGLE = (short)2;
        public const short ST_ELLIPSE = (short)3;
        public const short ST_DIAMOND = (short)4;
        public const short ST_ISOCELESTRIANGLE = (short)5;
        public const short ST_RIGHTTRIANGLE = (short)6;
        public const short ST_PARALLELOGRAM = (short)7;
        public const short ST_TRAPEZOID = (short)8;
        public const short ST_HEXAGON = (short)9;
        public const short ST_OCTAGON = (short)10;
        public const short ST_PLUS = (short)11;
        public const short ST_STAR = (short)12;
        public const short ST_ARROW = (short)13;
        public const short ST_THICKARROW = (short)14;
        public const short ST_HOMEPLATE = (short)15;
        public const short ST_CUBE = (short)16;
        public const short ST_BALLOON = (short)17;
        public const short ST_SEAL = (short)18;
        public const short ST_ARC = (short)19;
        public const short ST_LINE = (short)20;
        public const short ST_PLAQUE = (short)21;
        public const short ST_CAN = (short)22;
        public const short ST_DONUT = (short)23;
        public const short ST_TEXTSIMPLE = (short)24;
        public const short ST_TEXTOCTAGON = (short)25;
        public const short ST_TEXTHEXAGON = (short)26;
        public const short ST_TEXTCURVE = (short)27;
        public const short ST_TEXTWAVE = (short)28;
        public const short ST_TEXTRING = (short)29;
        public const short ST_TEXTONCURVE = (short)30;
        public const short ST_TEXTONRING = (short)31;
        public const short ST_STRAIGHTCONNECTOR1 = (short)32;
        public const short ST_BENTCONNECTOR2 = (short)33;
        public const short ST_BENTCONNECTOR3 = (short)34;
        public const short ST_BENTCONNECTOR4 = (short)35;
        public const short ST_BENTCONNECTOR5 = (short)36;
        public const short ST_CURVEDCONNECTOR2 = (short)37;
        public const short ST_CURVEDCONNECTOR3 = (short)38;
        public const short ST_CURVEDCONNECTOR4 = (short)39;
        public const short ST_CURVEDCONNECTOR5 = (short)40;
        public const short ST_CALLOUT1 = (short)41;
        public const short ST_CALLOUT2 = (short)42;
        public const short ST_CALLOUT3 = (short)43;
        public const short ST_ACCENTCALLOUT1 = (short)44;
        public const short ST_ACCENTCALLOUT2 = (short)45;
        public const short ST_ACCENTCALLOUT3 = (short)46;
        public const short ST_BORDERCALLOUT1 = (short)47;
        public const short ST_BORDERCALLOUT2 = (short)48;
        public const short ST_BORDERCALLOUT3 = (short)49;
        public const short ST_ACCENTBORDERCALLOUT1 = (short)50;
        public const short ST_ACCENTBORDERCALLOUT2 = (short)51;
        public const short ST_ACCENTBORDERCALLOUT3 = (short)52;
        public const short ST_RIBBON = (short)53;
        public const short ST_RIBBON2 = (short)54;
        public const short ST_CHEVRON = (short)55;
        public const short ST_PENTAGON = (short)56;
        public const short ST_NOSMOKING = (short)57;
        public const short ST_SEAL8 = (short)58;
        public const short ST_SEAL16 = (short)59;
        public const short ST_SEAL32 = (short)60;
        public const short ST_WEDGERECTCALLOUT = (short)61;
        public const short ST_WEDGERRECTCALLOUT = (short)62;
        public const short ST_WEDGEELLIPSECALLOUT = (short)63;
        public const short ST_WAVE = (short)64;
        public const short ST_FOLDEDCORNER = (short)65;
        public const short ST_LEFTARROW = (short)66;
        public const short ST_DOWNARROW = (short)67;
        public const short ST_UPARROW = (short)68;
        public const short ST_LEFTRIGHTARROW = (short)69;
        public const short ST_UPDOWNARROW = (short)70;
        public const short ST_IRREGULARSEAL1 = (short)71;
        public const short ST_IRREGULARSEAL2 = (short)72;
        public const short ST_LIGHTNINGBOLT = (short)73;
        public const short ST_HEART = (short)74;
        public const short ST_PICTUREFRAME = (short)75;
        public const short ST_QUADARROW = (short)76;
        public const short ST_LEFTARROWCALLOUT = (short)77;
        public const short ST_RIGHTARROWCALLOUT = (short)78;
        public const short ST_UPARROWCALLOUT = (short)79;
        public const short ST_DOWNARROWCALLOUT = (short)80;
        public const short ST_LEFTRIGHTARROWCALLOUT = (short)81;
        public const short ST_UPDOWNARROWCALLOUT = (short)82;
        public const short ST_QUADARROWCALLOUT = (short)83;
        public const short ST_BEVEL = (short)84;
        public const short ST_LEFTBRACKET = (short)85;
        public const short ST_RIGHTBRACKET = (short)86;
        public const short ST_LEFTBRACE = (short)87;
        public const short ST_RIGHTBRACE = (short)88;
        public const short ST_LEFTUPARROW = (short)89;
        public const short ST_BENTUPARROW = (short)90;
        public const short ST_BENTARROW = (short)91;
        public const short ST_SEAL24 = (short)92;
        public const short ST_STRIPEDRIGHTARROW = (short)93;
        public const short ST_NOTCHEDRIGHTARROW = (short)94;
        public const short ST_BLOCKARC = (short)95;
        public const short ST_SMILEYFACE = (short)96;
        public const short ST_VERTICALSCROLL = (short)97;
        public const short ST_HORIZONTALSCROLL = (short)98;
        public const short ST_CIRCULARARROW = (short)99;
        public const short ST_NOTCHEDCIRCULARARROW = (short)100;
        public const short ST_UTURNARROW = (short)101;
        public const short ST_CURVEDRIGHTARROW = (short)102;
        public const short ST_CURVEDLEFTARROW = (short)103;
        public const short ST_CURVEDUPARROW = (short)104;
        public const short ST_CURVEDDOWNARROW = (short)105;
        public const short ST_CLOUDCALLOUT = (short)106;
        public const short ST_ELLIPSERIBBON = (short)107;
        public const short ST_ELLIPSERIBBON2 = (short)108;
        public const short ST_FLOWCHARTProcess = (short)109;
        public const short ST_FLOWCHARTDECISION = (short)110;
        public const short ST_FLOWCHARTINPUTOUTPUT = (short)111;
        public const short ST_FLOWCHARTPREDEFINEDProcess = (short)112;
        public const short ST_FLOWCHARTINTERNALSTORAGE = (short)113;
        public const short ST_FLOWCHARTDOCUMENT = (short)114;
        public const short ST_FLOWCHARTMULTIDOCUMENT = (short)115;
        public const short ST_FLOWCHARTTERMINATOR = (short)116;
        public const short ST_FLOWCHARTPREPARATION = (short)117;
        public const short ST_FLOWCHARTMANUALINPUT = (short)118;
        public const short ST_FLOWCHARTMANUALOPERATION = (short)119;
        public const short ST_FLOWCHARTCONNECTOR = (short)120;
        public const short ST_FLOWCHARTPUNCHEDCARD = (short)121;
        public const short ST_FLOWCHARTPUNCHEDTAPE = (short)122;
        public const short ST_FLOWCHARTSUMMINGJUNCTION = (short)123;
        public const short ST_FLOWCHARTOR = (short)124;
        public const short ST_FLOWCHARTCOLLATE = (short)125;
        public const short ST_FLOWCHARTSORT = (short)126;
        public const short ST_FLOWCHARTEXTRACT = (short)127;
        public const short ST_FLOWCHARTMERGE = (short)128;
        public const short ST_FLOWCHARTOFFLINESTORAGE = (short)129;
        public const short ST_FLOWCHARTONLINESTORAGE = (short)130;
        public const short ST_FLOWCHARTMAGNETICTAPE = (short)131;
        public const short ST_FLOWCHARTMAGNETICDISK = (short)132;
        public const short ST_FLOWCHARTMAGNETICDRUM = (short)133;
        public const short ST_FLOWCHARTDISPLAY = (short)134;
        public const short ST_FLOWCHARTDELAY = (short)135;
        public const short ST_TEXTPLAINTEXT = (short)136;
        public const short ST_TEXTSTOP = (short)137;
        public const short ST_TEXTTRIANGLE = (short)138;
        public const short ST_TEXTTRIANGLEINVERTED = (short)139;
        public const short ST_TEXTCHEVRON = (short)140;
        public const short ST_TEXTCHEVRONINVERTED = (short)141;
        public const short ST_TEXTRINGINSIDE = (short)142;
        public const short ST_TEXTRINGOUTSIDE = (short)143;
        public const short ST_TEXTARCHUPCURVE = (short)144;
        public const short ST_TEXTARCHDOWNCURVE = (short)145;
        public const short ST_TEXTCIRCLECURVE = (short)146;
        public const short ST_TEXTBUTTONCURVE = (short)147;
        public const short ST_TEXTARCHUPPOUR = (short)148;
        public const short ST_TEXTARCHDOWNPOUR = (short)149;
        public const short ST_TEXTCIRCLEPOUR = (short)150;
        public const short ST_TEXTBUTTONPOUR = (short)151;
        public const short ST_TEXTCURVEUP = (short)152;
        public const short ST_TEXTCURVEDOWN = (short)153;
        public const short ST_TEXTCASCADEUP = (short)154;
        public const short ST_TEXTCASCADEDOWN = (short)155;
        public const short ST_TEXTWAVE1 = (short)156;
        public const short ST_TEXTWAVE2 = (short)157;
        public const short ST_TEXTWAVE3 = (short)158;
        public const short ST_TEXTWAVE4 = (short)159;
        public const short ST_TEXTINFLATE = (short)160;
        public const short ST_TEXTDEFLATE = (short)161;
        public const short ST_TEXTINFLATEBOTTOM = (short)162;
        public const short ST_TEXTDEFLATEBOTTOM = (short)163;
        public const short ST_TEXTINFLATETOP = (short)164;
        public const short ST_TEXTDEFLATETOP = (short)165;
        public const short ST_TEXTDEFLATEINFLATE = (short)166;
        public const short ST_TEXTDEFLATEINFLATEDEFLATE = (short)167;
        public const short ST_TEXTFADERIGHT = (short)168;
        public const short ST_TEXTFADELEFT = (short)169;
        public const short ST_TEXTFADEUP = (short)170;
        public const short ST_TEXTFADEDOWN = (short)171;
        public const short ST_TEXTSLANTUP = (short)172;
        public const short ST_TEXTSLANTDOWN = (short)173;
        public const short ST_TEXTCANUP = (short)174;
        public const short ST_TEXTCANDOWN = (short)175;
        public const short ST_FLOWCHARTALTERNATEProcess = (short)176;
        public const short ST_FLOWCHARTOFFPAGECONNECTOR = (short)177;
        public const short ST_CALLOUT90 = (short)178;
        public const short ST_ACCENTCALLOUT90 = (short)179;
        public const short ST_BORDERCALLOUT90 = (short)180;
        public const short ST_ACCENTBORDERCALLOUT90 = (short)181;
        public const short ST_LEFTRIGHTUPARROW = (short)182;
        public const short ST_SUN = (short)183;
        public const short ST_MOON = (short)184;
        public const short ST_BRACKETPAIR = (short)185;
        public const short ST_BRACEPAIR = (short)186;
        public const short ST_SEAL4 = (short)187;
        public const short ST_DOUBLEWAVE = (short)188;
        public const short ST_ACTIONBUTTONBLANK = (short)189;
        public const short ST_ACTIONBUTTONHOME = (short)190;
        public const short ST_ACTIONBUTTONHELP = (short)191;
        public const short ST_ACTIONBUTTONINFORMATION = (short)192;
        public const short ST_ACTIONBUTTONFORWARDNEXT = (short)193;
        public const short ST_ACTIONBUTTONBACKPREVIOUS = (short)194;
        public const short ST_ACTIONBUTTONEND = (short)195;
        public const short ST_ACTIONBUTTONBEGINNING = (short)196;
        public const short ST_ACTIONBUTTONRETURN = (short)197;
        public const short ST_ACTIONBUTTONDOCUMENT = (short)198;
        public const short ST_ACTIONBUTTONSOUND = (short)199;
        public const short ST_ACTIONBUTTONMOVIE = (short)200;
        public const short ST_HOSTCONTROL = (short)201;
        public const short ST_TEXTBOX = (short)202;
        public const short ST_NIL = (short)0x0FFF;

        protected HSSFPatriarch patriarch;

        /** Maps shape container objects to their OBJ records */
        private Dictionary<EscherRecord, Record> shapeToObj = new Dictionary<EscherRecord, Record>();
        private DrawingManager2 drawingManager;
        //private short drawingGroupId;

        /**
         * list of "tail" records that need to be Serialized after all drawing Group records
         */
        internal Dictionary<int, NoteRecord> tailRec = new Dictionary<int, NoteRecord>();

        public EscherAggregate(bool createDefaultTree)
        {
            if (createDefaultTree)
            {
                BuildBaseTree();
            }
        }

        /**
         * @return  Returns the current sid.
         */
        public override short Sid
        {
            get { return sid; }
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
                result.Append(escherRecord.ToString() + nl);
            }
            result.Append("[/").Append(RecordName).Append(']' + nl);

            return result.ToString();
        }
        /**
         * Calculates the xml representation of this record.  This is
         * simply a dump of all the records.
         * @param tab - string which must be added before each line (used by default '\t')
         * @return xml representation of the all aggregated records
         */
        public String ToXml(String tab)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(tab).Append("<").Append(RecordName).Append(">\n");
            for (IEnumerator iterator = EscherRecords.GetEnumerator(); iterator.MoveNext(); )
            {
                EscherRecord escherRecord = (EscherRecord)iterator.Current;
                builder.Append(escherRecord.ToXml(tab + "\t"));
            }
            builder.Append(tab).Append("</").Append(RecordName).Append(">\n");
            return builder.ToString();
        }

        /**
         * @param sid - record sid we want to check if it belongs to drawing layer
         * @return true if record is instance of DrawingRecord or ContinueRecord or ObjRecord or TextObjRecord
         */
        private static bool IsDrawingLayerRecord(short sid)
        {
            return sid == DrawingRecord.sid ||
                    sid == ContinueRecord.sid ||
                    sid == ObjRecord.sid ||
                    sid == TextObjectRecord.sid;
        }

        /**
         * Collapses the drawing records into an aggregate.
         * read Drawing, Obj, TxtObj, Note and Continue records into single byte array,
         * create Escher tree from byte array, create map &lt;EscherRecord, Record&gt;
         *
         * @param records - list of all records inside sheet
         * @param locFirstDrawingRecord - location of the first DrawingRecord inside sheet
         * @return new EscherAggregate create from all aggregated records which belong to drawing layer
         */
        public static EscherAggregate CreateAggregate(List<RecordBase> records, int locFirstDrawingRecord)
        {
            // Keep track of any shape records Created so we can match them back to the object id's.
            // Textbox objects are also treated as shape objects.
            List<EscherRecord> shapeRecords = new List<EscherRecord>();
            IEscherRecordFactory recordFactory = new CustomEscherRecordFactory(shapeRecords);

            // Create one big buffer
            MemoryStream stream = new MemoryStream();
            EscherAggregate agg = new EscherAggregate(false);
            int loc = locFirstDrawingRecord;
            while (loc + 1 < records.Count
                    && (IsDrawingLayerRecord(GetSid(records, loc))))
            {
                try
                {
                    if (!(GetSid(records, loc) == DrawingRecord.sid || GetSid(records, loc) == ContinueRecord.sid))
                    {
                        loc++;
                        continue;
                    }
                    if (GetSid(records, loc) == DrawingRecord.sid)
                    {
                        byte[] data = ((DrawingRecord)records[loc]).Data;
                        stream.Write(data, 0, data.Length);
                    }
                    else
                    {
                        byte[] data = ((ContinueRecord)records[loc]).Data;
                        stream.Write(data, 0, data.Length);
                    }
                }
                catch (IOException e)
                {
                    throw new RuntimeException("Couldn't get data from drawing/continue records", e);
                }
                loc++;
            }

            // Decode the shapes
            //        agg.escherRecords = new ArrayList();
            int pos = 0;
            byte[] buffer = stream.ToArray();
            while (pos < buffer.Length)
            {
                EscherRecord r = recordFactory.CreateRecord(buffer, pos);
                int bytesRead = r.FillFields(buffer, pos, recordFactory);
                agg.AddEscherRecord(r);
                pos += bytesRead;
            }

            // Associate the object records with the shapes
            loc = locFirstDrawingRecord + 1;
            int shapeIndex = 0;

            while (loc < records.Count && IsDrawingLayerRecord(GetSid(records, loc)))
            {
                if (!IsObjectRecord(records, loc))
                {
                    loc++;
                    continue;
                }
                Record objRecord = (Record)records[loc];
                agg.shapeToObj[shapeRecords[shapeIndex++]] = objRecord;
                loc++;
            }
            // any NoteRecords that follow the drawing block must be aggregated and and saved in the tailRec collection
            //put noterecord into tailsRec
            while (loc < records.Count)
            {
                if (GetSid(records, loc) == NoteRecord.sid)
                {
                    NoteRecord r = (NoteRecord)records[(loc)];
                    agg.tailRec[r.ShapeId] = r;
                }
                else
                {
                    break;
                }
                loc++;
            }
            int locLastDrawingRecord = loc;
            //records.SubList(locFirstDrawingRecord, locLastDrawingRecord).Clear();
            records.RemoveRange(locFirstDrawingRecord, locLastDrawingRecord - locFirstDrawingRecord);
            records.Insert(locFirstDrawingRecord, agg);
            return agg;

        }
        /**
         * Serializes this aggregate to a byte array.  Since this Is an aggregate
         * record it will effectively Serialize the aggregated records.
         *
         * @param offset    The offset into the start of the array.
         * @param data      The byte array to Serialize to.
         * @return          The number of bytes Serialized.
         */
        public override int Serialize(int offset, byte[] data)
        {

            // Determine buffer size
            List<EscherRecord> records = EscherRecords;
            int size = GetEscherRecordSize(records);
            byte[] buffer = new byte[size];


            // Serialize escher records into one big data structure and keep note of ending offsets.
            List<int> spEndingOffsets = new List<int>();
            List<EscherRecord> shapes = new List<EscherRecord>();
            int pos = 0;
            foreach (EscherRecord e in records)
            {
                pos += e.Serialize(pos, buffer, new SerializationListener(spEndingOffsets, shapes, e));
            }
            // todo: fix this
            shapes.Insert(0, null);
            spEndingOffsets.Insert(0, 0);

            // Split escher records into Separate MSODRAWING and OBJ, TXO records.  (We don't break on
            // the first one because it's the patriach).
            pos = offset;
            int writtenEscherBytes = 0;
            int i;
            for (i = 1; i < shapes.Count; i++)
            {
                int endOffset = (int)spEndingOffsets[i] - 1;
                int startOffset;
                if (i == 1)
                    startOffset = 0;
                else
                    startOffset = (int)spEndingOffsets[i - 1];

                byte[] drawingData = new byte[endOffset - startOffset + 1];
                Array.Copy(buffer, startOffset, drawingData, 0, drawingData.Length);
                pos += WriteDataIntoDrawingRecord(drawingData, writtenEscherBytes, pos, data, i);

                writtenEscherBytes += drawingData.Length;

                // Write the matching OBJ record
                Record obj = (Record)shapeToObj[shapes[i]];
                pos += obj.Serialize(pos, data);

                if (i == shapes.Count - 1 && endOffset < buffer.Length - 1)
                {
                    drawingData = new byte[buffer.Length - endOffset - 1];
                    Array.Copy(buffer, endOffset + 1, drawingData, 0, drawingData.Length);
                    pos += WriteDataIntoDrawingRecord(drawingData, writtenEscherBytes, pos, data, i);
                }

            }
            if ((pos - offset) < buffer.Length - 1)
            {
                byte[] drawingData = new byte[buffer.Length - (pos - offset)];
                System.Array.Copy(buffer, (pos - offset), drawingData, 0, drawingData.Length);
                pos += WriteDataIntoDrawingRecord(drawingData, writtenEscherBytes, pos, data, i);
            }
            // Write records that need to be Serialized after all drawing Group records
            foreach (Record rec in tailRec.Values)
            {
                pos += rec.Serialize(pos, data);
            }
            int bytesWritten = pos - offset;
            if (bytesWritten != RecordSize)
                throw new RecordFormatException(bytesWritten + " bytes written but RecordSize reports " + RecordSize);
            return bytesWritten;
        }

        /**
         * @param drawingData - escher records saved into single byte array
         * @param writtenEscherBytes - count of bytes already saved into drawing records (we should know it to decide create
         *                           drawing or continue record)
         * @param pos current position of data array
         * @param data - array of bytes where drawing records must be serialized
         * @param i - number of shape, saved into data array
         * @return offset of data array after serialization
         */
        private int WriteDataIntoDrawingRecord(byte[] drawingData, int writtenEscherBytes, int pos, byte[] data, int i)
        {
            int temp = 0;
            //First record in drawing layer MUST be DrawingRecord
            if (writtenEscherBytes + drawingData.Length > RecordInputStream.MAX_RECORD_DATA_SIZE && i != 1)
            {
                for (int j = 0; j < drawingData.Length; j += RecordInputStream.MAX_RECORD_DATA_SIZE)
                {
                    byte[] buf = new byte[Math.Min(RecordInputStream.MAX_RECORD_DATA_SIZE, drawingData.Length - j)];
                    System.Array.Copy(drawingData, j, buf, 0, Math.Min(RecordInputStream.MAX_RECORD_DATA_SIZE, drawingData.Length - j));
                    ContinueRecord drawing = new ContinueRecord(buf);
                    temp += drawing.Serialize(pos + temp, data);
                }
            }
            else
            {
                for (int j = 0; j < drawingData.Length; j += RecordInputStream.MAX_RECORD_DATA_SIZE)
                {
                    if (j == 0)
                    {
                        DrawingRecord drawing = new DrawingRecord();
                        byte[] buf = new byte[Math.Min(RecordInputStream.MAX_RECORD_DATA_SIZE, drawingData.Length - j)];
                        System.Array.Copy(drawingData, j, buf, 0, Math.Min(RecordInputStream.MAX_RECORD_DATA_SIZE, drawingData.Length - j));
                        drawing.Data = (buf);
                        temp += drawing.Serialize(pos + temp, data);
                    }
                    else
                    {
                        byte[] buf = new byte[Math.Min(RecordInputStream.MAX_RECORD_DATA_SIZE, drawingData.Length - j)];
                        System.Array.Copy(drawingData, j, buf, 0, Math.Min(RecordInputStream.MAX_RECORD_DATA_SIZE, drawingData.Length - j));
                        ContinueRecord drawing = new ContinueRecord(buf);
                        temp += drawing.Serialize(pos + temp, data);
                    }
                }
            }
            return temp;
        }

        /**
         * How many bytes do the raw escher records contain.
         *
         * @param records List of escher records
         * @return the number of bytes
         */
        private int GetEscherRecordSize(List<EscherRecord> records)
        {
            int size = 0;
            foreach (EscherRecord record in records)
            {
                size += record.RecordSize;
            }
            return size;
        }
        /**
         * @param records list of records to look into
         * @param loc - location of the record which sid must be returned
         * @return sid of the record with selected location
         */
        private static short GetSid(List<RecordBase> records, int loc)
        {
            RecordBase record = records[(loc)];
            if (record is Record)
            {
                return ((Record)record).Sid;
            }
            else
            {
                // Aggregates don't have a sid
                // We could step into them, but for these needs we don't care
                return -1;
            }
        }


        /**
     * @return record size, including header size of obj, text, note, drawing, continue records
     */
        public override int RecordSize
        {
            get
            {
                // To determine size of aggregate record we have to know size of each DrawingRecord because if DrawingRecord
                // is split into several continue records we have to add header size to total EscherAggregate size
                int continueRecordsHeadersSize = 0;
                // Determine buffer size
                List<EscherRecord> records = EscherRecords;
                int rawEscherSize = GetEscherRecordSize(records);
                byte[] buffer = new byte[rawEscherSize];
                List<int> spEndingOffsets = new List<int>();
                int pos = 0;
                foreach (EscherRecord e in records)
                {
                    pos += e.Serialize(pos, buffer, new RecordSizeListener(spEndingOffsets, e));
                }
                spEndingOffsets.Insert(0, 0);

                for (int i = 1; i < spEndingOffsets.Count; i++)
                {
                    if (i == spEndingOffsets.Count - 1 && spEndingOffsets[(i)] < pos)
                    {
                        continueRecordsHeadersSize += 4;
                    }
                    if (spEndingOffsets[(i)] - spEndingOffsets[(i - 1)] <= RecordInputStream.MAX_RECORD_DATA_SIZE)
                    {
                        continue;
                    }
                    continueRecordsHeadersSize += ((spEndingOffsets[(i)] - spEndingOffsets[(i - 1)]) / RecordInputStream.MAX_RECORD_DATA_SIZE) * 4;
                }

                int drawingRecordSize = rawEscherSize + (shapeToObj.Count) * 4;
                if (rawEscherSize != 0 && spEndingOffsets.Count == 1/*EMPTY*/)
                {
                    continueRecordsHeadersSize += 4;
                }
                int objRecordSize = 0;
                foreach (Record r in shapeToObj.Values)
                {
                    objRecordSize += r.RecordSize;
                }
                int tailRecordSize = 0;
                foreach (NoteRecord noteRecord in tailRec.Values)
                {
                    tailRecordSize += noteRecord.RecordSize;
                }
                return drawingRecordSize + objRecordSize + tailRecordSize + continueRecordsHeadersSize;
            }
        }

        /**
         * create base tree with such structure:
         * EscherDgContainer
         * -EscherSpgrContainer
         * --EscherSpContainer
         * ---EscherSpRecord
         * ---EscherSpgrRecord
         * ---EscherSpRecord
         * -EscherDgRecord
         *
         * id of DgRecord and SpRecord are empty and must be set later by HSSFPatriarch
         */
        private void BuildBaseTree()
        {
            EscherContainerRecord dgContainer = new EscherContainerRecord();
            EscherContainerRecord spgrContainer = new EscherContainerRecord();
            EscherContainerRecord spContainer1 = new EscherContainerRecord();
            EscherSpgrRecord spgr = new EscherSpgrRecord();
            EscherSpRecord sp1 = new EscherSpRecord();
            dgContainer.RecordId = (EscherContainerRecord.DG_CONTAINER);
            dgContainer.Options = ((short)0x000F);
            EscherDgRecord dg = new EscherDgRecord();
            dg.RecordId = EscherDgRecord.RECORD_ID;
            short dgId = 1;
            dg.Options = ((short)(dgId << 4));
            dg.NumShapes = (0);
            dg.LastMSOSPID=(1024);
            spgrContainer.RecordId=(EscherContainerRecord.SPGR_CONTAINER);
            spgrContainer.Options=((short)0x000F);
            spContainer1.RecordId=(EscherContainerRecord.SP_CONTAINER);
            spContainer1.Options=((short)0x000F);
            spgr.RecordId=(EscherSpgrRecord.RECORD_ID);
            spgr.Options=((short)0x0001);    // version
            spgr.RectX1=(0);
            spgr.RectY1=(0);
            spgr.RectX2=(1023);
            spgr.RectY2=(255);
            sp1.RecordId=(EscherSpRecord.RECORD_ID);

            sp1.Options=((short)0x0002);
            sp1.Version=((short)0x2);
            sp1.ShapeId=(-1);
            sp1.Flags=(EscherSpRecord.FLAG_GROUP | EscherSpRecord.FLAG_PATRIARCH);
            dgContainer.AddChildRecord(dg);
            dgContainer.AddChildRecord(spgrContainer);
            spgrContainer.AddChildRecord(spContainer1);
            spContainer1.AddChildRecord(spgr);
            spContainer1.AddChildRecord(sp1);
            AddEscherRecord(dgContainer);
        }

        public EscherAggregate(DrawingManager2 drawingManager)
        {
            this.drawingManager = drawingManager;
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

        

        

        internal class CustomEscherRecordFactory : DefaultEscherRecordFactory
        {
            List<EscherRecord> shapeRecords;
            public CustomEscherRecordFactory(List<EscherRecord> shapeRecords)
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

        public HSSFPatriarch Patriarch
        {
            get { return patriarch; }
            set { this.patriarch = value; }
        }

        /**
         * Converts the Records into UserModel
         *  objects on the bound HSSFPatriarch
         */
        //public void ConvertRecordsToUserModel()
        //{
        //    if (patriarch == null)
        //    {
        //        throw new InvalidOperationException("Must call SetPatriarch() first");
        //    }

        //    // The top level container ought to have
        //    //  the DgRecord and the container of one container
        //    //  per shape Group (patriach overall first)
        //    EscherContainerRecord topContainer =
        //        (EscherContainerRecord)GetEscherContainer();
        //    if (topContainer == null)
        //    {
        //        return;
        //    }
        //    topContainer = (EscherContainerRecord)
        //        topContainer.ChildContainers[0];

        //    IList<EscherContainerRecord> tcc = topContainer.ChildContainers;
        //    if (tcc.Count == 0)
        //    {
        //        throw new InvalidOperationException("No child escher containers at the point that should hold the patriach data, and one container per top level shape!");
        //    }

        //    // First up, Get the patriach position
        //    // This Is in the first EscherSpgrRecord, in
        //    //  the first container, with a EscherSRecord too
        //    EscherContainerRecord patriachContainer =
        //        (EscherContainerRecord)tcc[0];
        //    EscherSpgrRecord spgr = null;
        //    for (IEnumerator it = patriachContainer.ChildRecords.GetEnumerator(); it.MoveNext(); )
        //    {
        //        EscherRecord r = (EscherRecord)it.Current;
        //        if (r is EscherSpgrRecord)
        //        {
        //            spgr = (EscherSpgrRecord)r;
        //            break;
        //        }
        //    }
        //    if (spgr != null)
        //    {
        //        patriarch.SetCoordinates(
        //                spgr.RectX1, spgr.RectY1,
        //                spgr.RectX2, spgr.RectY2
        //        );
        //    }

        //    // Now Process the containers for each Group
        //    //  and objects
        //    for (int i = 1; i < tcc.Count; i++)
        //    {
        //        EscherContainerRecord shapeContainer =
        //            (EscherContainerRecord)tcc[i];
        //        //Console.Error.WriteLine("\n\n*****\n\n");
        //        //Console.Error.WriteLine(shapeContainer);

        //        // Could be a Group, or a base object
        //        if (shapeContainer.RecordId == EscherContainerRecord.SPGR_CONTAINER)
        //        {
        //            if(shapeContainer.ChildRecords.Count>0)
        //            {
        //                // Group
        //                HSSFShapeGroup group =
        //                    new HSSFShapeGroup(null, new HSSFClientAnchor());
        //                patriarch.Children.Add(group);

        //                EscherContainerRecord groupContainer =
        //                    (EscherContainerRecord)shapeContainer.GetChild(0);
        //                ConvertRecordsToUserModel(groupContainer, group);
        //            }
        //        }
        //        else if (shapeContainer.RecordId == EscherContainerRecord.SP_CONTAINER)
        //        {
        //            EscherSpRecord spRecord = (EscherSpRecord)shapeContainer.GetChildById(EscherSpRecord.RECORD_ID);
        //            int type = spRecord.ShapeType;

        //            switch (type)
        //            {
        //                case ST_TEXTBOX:
        //                    HSSFSimpleShape box;

        //                    TextObjectRecord textrec = (TextObjectRecord)shapeToObj[GetEscherChild(shapeContainer, EscherTextboxRecord.RECORD_ID)];
        //                     EscherClientAnchorRecord anchorRecord1 = (EscherClientAnchorRecord)GetEscherChild(shapeContainer, EscherClientAnchorRecord.RECORD_ID);
        //                        HSSFClientAnchor anchor1 = new HSSFClientAnchor();
        //                        anchor1.Col1 = anchorRecord1.Col1;
        //                        anchor1.Col2 = anchorRecord1.Col2;
        //                        anchor1.Dx1 = anchorRecord1.Dx1;
        //                        anchor1.Dx2 = anchorRecord1.Dx2;
        //                        anchor1.Dy1 = anchorRecord1.Dy1;
        //                        anchor1.Dy2 = anchorRecord1.Dy2;
        //                        anchor1.Row1 = anchorRecord1.Row1;
        //                        anchor1.Row2 = anchorRecord1.Row2;
        //                    if (tailRec.Count>=i && tailRec[i-1] is NoteRecord)
        //                    {
        //                        NoteRecord noterec=(NoteRecord)tailRec[i - 1];
                                
        //                        // comment
        //                        box =
        //                            new HSSFComment(null, anchor1);
        //                        HSSFComment comment=(HSSFComment)box;
        //                        comment.Author = noterec.Author;
        //                        comment.Row = noterec.Row;
        //                        comment.Column = noterec.Column;
        //                        comment.Visible = (noterec.Flags == NoteRecord.NOTE_VISIBLE);
        //                        comment.String = textrec.Str;                                
        //                    }
        //                    else
        //                    {
        //                        // TextBox
        //                        box =
        //                            new HSSFTextbox(null, anchor1);
        //                        ((HSSFTextbox)box).String = textrec.Str;  
        //                    }
        //                    patriarch.AddShape(box);
        //                    ConvertRecordsToUserModel(shapeContainer, box);
        //                    break;
        //                case ST_PICTUREFRAME:
        //                    // Duplicated from
        //                    // org.apache.poi.hslf.model.Picture.getPictureIndex()
        //                    EscherOptRecord opt = (EscherOptRecord)GetEscherChild(shapeContainer, EscherOptRecord.RECORD_ID);
        //                    EscherSimpleProperty prop = (EscherSimpleProperty)opt.Lookup(EscherProperties.BLIP__BLIPTODISPLAY);
        //                    if (prop != null)
        //                    {
        //                        int pictureIndex = prop.PropertyValue;
        //                        EscherClientAnchorRecord anchorRecord = (EscherClientAnchorRecord)GetEscherChild(shapeContainer, EscherClientAnchorRecord.RECORD_ID);
        //                        HSSFClientAnchor anchor = new HSSFClientAnchor();
        //                        anchor.Col1 = anchorRecord.Col1;
        //                        anchor.Col2 = anchorRecord.Col2;
        //                        anchor.Dx1 = anchorRecord.Dx1;
        //                        anchor.Dx2 = anchorRecord.Dx2;
        //                        anchor.Dy1 = anchorRecord.Dy1;
        //                        anchor.Dy2 = anchorRecord.Dy2;
        //                        anchor.Row1 = anchorRecord.Row1;
        //                        anchor.Row2 = anchorRecord.Row2;
        //                        HSSFPicture picture = new HSSFPicture(null, anchor);
        //                        picture.PictureIndex = pictureIndex;
        //                        patriarch.AddShape(picture);

        //                    }
        //                    break;
        //            }


        //        }
        //        else
        //        {
        //            // Base level
        //            ConvertRecordsToUserModel(shapeContainer, patriarch);
        //        }
        //    }

        //    // Now, clear any trace of what records make up
        //    //  the patriarch
        //    // Otherwise, everything will go horribly wrong
        //    //  when we try to Write out again....
        //    //    	clearEscherRecords();
        //    drawingManager.GetDgg().FileIdClusters=new EscherDggRecord.FileIdCluster[0];

        //    // TODO: Support Converting our records
        //    //  back into shapes
        //    log.Log(POILogger.WARN, "Not Processing objects into Patriarch!");
        //}

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


        //private void ConvertRecordsToUserModel(EscherContainerRecord shapeContainer, Object model)
        //{
        //    for (IEnumerator it = shapeContainer.ChildRecords.GetEnumerator(); it.MoveNext(); )
        //    {
        //        EscherRecord r = (EscherRecord)it.Current;
        //        if (r is EscherSpgrRecord)
        //        {
        //            // This may be overriden by a later EscherClientAnchorRecord
        //            EscherSpgrRecord spgr = (EscherSpgrRecord)r;

        //            if (model is HSSFShapeGroup)
        //            {
        //                HSSFShapeGroup g = (HSSFShapeGroup)model;
        //                g.SetCoordinates(
        //                        spgr.RectX1, spgr.RectY1,
        //                        spgr.RectX2, spgr.RectY2
        //                );
        //            }
        //            else
        //            {
        //                throw new InvalidOperationException("Got top level anchor but not Processing a Group");
        //            }
        //        }
        //        else if (r is EscherClientAnchorRecord)
        //        {
        //            EscherClientAnchorRecord car = (EscherClientAnchorRecord)r;

        //            if (model is HSSFShape)
        //            {
        //                HSSFShape g = (HSSFShape)model;
        //                g.Anchor.Dx1=car.Dx1;
        //                g.Anchor.Dx2=car.Dx2;
        //                g.Anchor.Dy1=car.Dy1;
        //                g.Anchor.Dy2=car.Dy2;
        //            }
        //            else
        //            {
        //                throw new InvalidOperationException("Got top level anchor but not Processing a Group or shape");
        //            }
        //        }
        //        else if (r is EscherTextboxRecord)
        //        {
        //            EscherTextboxRecord tbr = (EscherTextboxRecord)r;

        //            // Also need to Find the TextObjectRecord too
        //            // TODO
        //        }
        //        else if (r is EscherSpRecord)
        //        {
        //            // Use flags if needed
        //        }
        //        else if (r is EscherOptRecord)
        //        {
        //            // Use properties if needed
        //        }
        //        else
        //        {
        //            //Console.Error.WriteLine(r);
        //        }
        //    }
        //}

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

        private static bool IsObjectRecord(List<RecordBase> records, int loc)
        {
            return GetSid(records, loc) == ObjRecord.sid || GetSid(records, loc) == TextObjectRecord.sid;
        }
        

        //private void ConvertShapes(HSSFShapeContainer parent, EscherContainerRecord escherParent, Hashtable shapeToObj)
        //{
        //    if (escherParent == null) throw new ArgumentException("Parent record required");

        //    IList<HSSFShape> shapes = parent.Children;
        //    for (IEnumerator iterator = shapes.GetEnumerator(); iterator.MoveNext(); )
        //    {
        //        HSSFShape shape = (HSSFShape)iterator.Current;
        //        if (shape is HSSFShapeGroup)
        //        {
        //            ConvertGroup((HSSFShapeGroup)shape, escherParent, shapeToObj);
        //        }
        //        else
        //        {
        //            AbstractShape shapeModel = AbstractShape.CreateShape(
        //                    shape,
        //                    drawingManager.AllocateShapeId(drawingGroupId));
        //            shapeToObj[FindClientData(shapeModel.SpContainer)]=shapeModel.ObjRecord;
        //            if (shapeModel is TextboxShape)
        //            {
        //                EscherRecord escherTextbox = ((TextboxShape)shapeModel).EscherTextbox;
        //                shapeToObj[escherTextbox]=((TextboxShape)shapeModel).TextObjectRecord;
        //                //                    escherParent.AddChildRecord(escherTextbox);

        //                if (shapeModel is CommentShape)
        //                {
        //                    CommentShape comment = (CommentShape)shapeModel;
        //                    tailRec.Add(comment.NoteRecord.ShapeId,comment.NoteRecord);
        //                }
        //            }
        //            escherParent.AddChildRecord(shapeModel.SpContainer);
        //        }
        //    }
        //    //        drawingManager.newCluster( (short)1 );
        //    //        drawingManager.newCluster( (short)2 );

        //}

        //private void ConvertGroup(HSSFShapeGroup shape, EscherContainerRecord escherParent, Hashtable shapeToObj)
        //{
        //    EscherContainerRecord spgrContainer = new EscherContainerRecord();
        //    EscherContainerRecord spContainer = new EscherContainerRecord();
        //    EscherSpgrRecord spgr = new EscherSpgrRecord();
        //    EscherSpRecord sp = new EscherSpRecord();
        //    EscherOptRecord opt = new EscherOptRecord();
        //    EscherRecord anchor;
        //    EscherClientDataRecord clientData = new EscherClientDataRecord();

        //    spgrContainer.RecordId = EscherContainerRecord.SPGR_CONTAINER;
        //    spgrContainer.Options = (short)0x000F;
        //    spContainer.RecordId = EscherContainerRecord.SP_CONTAINER;
        //    spContainer.Options = (short)0x000F;
        //    spgr.RecordId = EscherSpgrRecord.RECORD_ID;
        //    spgr.Options = (short)0x0001;
        //    spgr.RectX1 = shape.X1;
        //    spgr.RectY1 = shape.Y1;
        //    spgr.RectX2 = shape.X2;
        //    spgr.RectY2 = shape.Y2;
        //    sp.RecordId = EscherSpRecord.RECORD_ID;
        //    sp.Options = (short)0x0002;
        //    int shapeId = drawingManager.AllocateShapeId(drawingGroupId);
        //    sp.ShapeId = shapeId;
        //    if (shape.Anchor is HSSFClientAnchor)
        //        sp.Flags = EscherSpRecord.FLAG_GROUP | EscherSpRecord.FLAG_HAVEANCHOR;
        //    else
        //        sp.Flags = EscherSpRecord.FLAG_GROUP | EscherSpRecord.FLAG_HAVEANCHOR | EscherSpRecord.FLAG_CHILD;
        //    opt.RecordId = EscherOptRecord.RECORD_ID;
        //    opt.Options = (short)0x0023;
        //    opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.PROTECTION__LOCKAGAINSTGROUPING, 0x00040004));
        //    opt.AddEscherProperty(new EscherBoolProperty(EscherProperties.GROUPSHAPE__PRINT, 0x00080000));

        //    anchor = ConvertAnchor.CreateAnchor(shape.Anchor);
        //    //        clientAnchor.Col1( ( (HSSFClientAnchor) shape.Anchor ).Col1 );
        //    //        clientAnchor.Row1( (short) ( (HSSFClientAnchor) shape.Anchor ).Row1 );
        //    //        clientAnchor.Dx1( (short) shape.Anchor.Dx1 );
        //    //        clientAnchor.Dy1( (short) shape.Anchor.Dy1 );
        //    //        clientAnchor.Col2( ( (HSSFClientAnchor) shape.Anchor ).Col2 );
        //    //        clientAnchor.Row2( (short) ( (HSSFClientAnchor) shape.Anchor ).Row2 );
        //    //        clientAnchor.Dx2( (short) shape.Anchor.Dx2 );
        //    //        clientAnchor.Dy2( (short) shape.Anchor.Dy2 );
        //    clientData.RecordId = (EscherClientDataRecord.RECORD_ID);
        //    clientData.Options = ((short)0x0000);

        //    spgrContainer.AddChildRecord(spContainer);
        //    spContainer.AddChildRecord(spgr);
        //    spContainer.AddChildRecord(sp);
        //    spContainer.AddChildRecord(opt);
        //    spContainer.AddChildRecord(anchor);
        //    spContainer.AddChildRecord(clientData);

        //    ObjRecord obj = new ObjRecord();
        //    CommonObjectDataSubRecord cmo = new CommonObjectDataSubRecord();
        //    cmo.ObjectType = CommonObjectType.GROUP;
        //    cmo.ObjectId = shapeId;
        //    cmo.IsLocked = true;
        //    cmo.IsPrintable = true;
        //    cmo.IsAutoFill = true;
        //    cmo.IsAutoline = true;
        //    GroupMarkerSubRecord gmo = new GroupMarkerSubRecord();
        //    EndSubRecord end = new EndSubRecord();
        //    obj.AddSubRecord(cmo);
        //    obj.AddSubRecord(gmo);
        //    obj.AddSubRecord(end);
        //    shapeToObj[clientData] = obj;

        //    escherParent.AddChildRecord(spgrContainer);

        //    ConvertShapes(shape, spgrContainer, shapeToObj);

        //}

        //private EscherRecord FindClientData(EscherContainerRecord spContainer)
        //{
        //    for (IEnumerator iterator = spContainer.ChildRecords.GetEnumerator(); iterator.MoveNext(); )
        //    {
        //        EscherRecord r = (EscherRecord)iterator.Current;
        //        if (r.RecordId == EscherClientDataRecord.RECORD_ID)
        //            return r;
        //    }
        //    throw new ArgumentException("Can not Find client data record");
        //}

        //private void ConvertPatriarch(HSSFPatriarch patriarch)
        //{
        //    EscherContainerRecord dgContainer = new EscherContainerRecord();
        //    EscherDgRecord dg;
        //    EscherContainerRecord spgrContainer = new EscherContainerRecord();
        //    EscherContainerRecord spContainer1 = new EscherContainerRecord();
        //    EscherSpgrRecord spgr = new EscherSpgrRecord();
        //    EscherSpRecord sp1 = new EscherSpRecord();

        //    dgContainer.RecordId=EscherContainerRecord.DG_CONTAINER;
        //    dgContainer.Options=(short)0x000F;
        //    dg = drawingManager.CreateDgRecord();
        //    drawingGroupId = dg.DrawingGroupId;
        //    //        dg.Options( (short) ( drawingId << 4 ) );
        //    //        dg.NumShapes( GetNumberOfShapes( patriarch ) );
        //    //        dg.LastMSOSPID( 0 );  // populated after all shape id's are assigned.
        //    spgrContainer.RecordId=EscherContainerRecord.SPGR_CONTAINER;
        //    spgrContainer.Options=(short)0x000F;
        //    spContainer1.RecordId=EscherContainerRecord.SP_CONTAINER;
        //    spContainer1.Options=(short)0x000F;
        //    spgr.RecordId=EscherSpgrRecord.RECORD_ID;
        //    spgr.Options=(short)0x0001;    // version
        //    spgr.RectX1=patriarch.X1;
        //    spgr.RectY1=patriarch.Y1;
        //    spgr.RectX2=patriarch.X2;
        //    spgr.RectY2=patriarch.Y2;
        //    sp1.RecordId=EscherSpRecord.RECORD_ID;
        //    sp1.Options=(short)0x0002;
        //    sp1.ShapeId=drawingManager.AllocateShapeId(dg.DrawingGroupId);
        //    sp1.Flags=EscherSpRecord.FLAG_GROUP | EscherSpRecord.FLAG_PATRIARCH;

        //    dgContainer.AddChildRecord(dg);
        //    dgContainer.AddChildRecord(spgrContainer);
        //    spgrContainer.AddChildRecord(spContainer1);
        //    spContainer1.AddChildRecord(spgr);
        //    spContainer1.AddChildRecord(sp1);

        //    AddEscherRecord(dgContainer);
        //}

        //private static short GetSid(IList records, int loc)
        //{
        //    return ((Record)records[loc]).Sid;
        //}
        /// <summary>
        /// Associates an escher record to an OBJ record or a TXO record.
        /// </summary>
        /// <param name="r">ClientData or Textbox record</param>
        /// <param name="objRecord">Obj or TextObj record</param>
        public void AssociateShapeToObjRecord(EscherRecord r, Record objRecord)
        {
            if(!shapeToObj.ContainsKey(r))
                shapeToObj.Add(r, objRecord);
            else
                shapeToObj[r]= objRecord;
        }
        /// <summary>
        /// Remove echerRecord and associated to it Obj or TextObj record
        /// </summary>
        /// <param name="rec">clientData or textbox record to be removed</param>
        public void RemoveShapeToObjRecord(EscherRecord rec)
        {
            shapeToObj.Remove(rec);
        }
        internal void SetDgId(short dgId)
        {
            EscherContainerRecord dgContainer = GetEscherContainer();
            EscherDgRecord dg = (EscherDgRecord)dgContainer.GetChildById(EscherDgRecord.RECORD_ID);
            dg.Options = (short)(dgId << 4);
        }
        internal void SetMainSpRecordId(int shapeId)
        {
            EscherContainerRecord dgContainer = GetEscherContainer();
            EscherContainerRecord spgrConatiner = (EscherContainerRecord)dgContainer.GetChildById(EscherContainerRecord.SPGR_CONTAINER);
            EscherContainerRecord spContainer = (EscherContainerRecord)spgrConatiner.GetChild(0);
            EscherSpRecord sp = (EscherSpRecord)spContainer.GetChildById(EscherSpRecord.RECORD_ID);
            sp.ShapeId = (shapeId);
        }

        public void RemoveTailRecord(NoteRecord note)
        {
            tailRec.Remove(note.ShapeId);
        }

        internal void AddTailRecord(NoteRecord note)
        {
            if (tailRec.ContainsKey(note.ShapeId))
                tailRec.Add(note.ShapeId, note);
            else
                tailRec[note.ShapeId] = note;
        }
        /**
     * @return unmodifiable copy of tail records. We need to access them when building shapes.
     *         Every HSSFComment shape has a link to a NoteRecord from the tailRec collection.
     */
        public Dictionary<int, NoteRecord> TailRecords { get { return tailRec; } }

        internal Dictionary<EscherRecord, Record> GetShapeToObjMapping()
        {
            return shapeToObj;
        }
        /**
     * @param obj - ObjRecord with id == NoteRecord.id
     * @return null if note record is not found else returns note record with id == obj.id
     */
        internal NoteRecord GetNoteRecordByObj(ObjRecord obj)
        {
            CommonObjectDataSubRecord cod = (CommonObjectDataSubRecord)obj.SubRecords[0];
            if (!tailRec.ContainsKey(cod.ObjectId))
                return null;
            return tailRec[(cod.ObjectId)];
        }
    }
}