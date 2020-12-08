/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 8, 2005
 *
 */

using NPOI.SS.Formula.Atp;

namespace NPOI.SS.Formula.Eval
{
    using System;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula.Function;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using Index = NPOI.SS.Formula.Functions.Index;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */

    public sealed class FunctionEval
    {
        private FunctionEval()
        {
            // no instances of this class
        }
        /**
         * Some function IDs that require special treatment
         */

        private class FunctionID
        {
            /** 1 */
            public const int IF = FunctionMetadataRegistry.FUNCTION_INDEX_IF;
            public const int SUM = FunctionMetadataRegistry.FUNCTION_INDEX_SUM;
            /** 78 */
            public const int OFFSET = 78;
            /** 100 */
            public const int CHOOSE = FunctionMetadataRegistry.FUNCTION_INDEX_CHOOSE;
            /** 148 */
            public const int INDIRECT = FunctionMetadataRegistry.FUNCTION_INDEX_INDIRECT;
            /** 255 */
            public const int EXTERNAL_FUNC = FunctionMetadataRegistry.FUNCTION_INDEX_EXTERNAL;
        }
        protected static Function[] functions = ProduceFunctions();

        // fix warning CS0169 "never used": private static Hashtable freeRefFunctionsByIdMap;
        private static FunctionMetadataRegistry _instance;

        private static FunctionMetadataRegistry GetInstance()
        {
            if (_instance == null)
            {
                _instance = FunctionMetadataReader.CreateRegistry();
            }
            return _instance;
        }


        public static Function GetBasicFunction(int functionIndex)
        {
            // check for 'free ref' functions first
            switch (functionIndex)
            {
                case FunctionID.INDIRECT:
                case FunctionID.EXTERNAL_FUNC:
                    return null;
            }
            // else - must be plain function
            Function result = functions[functionIndex];
            if (result == null)
            {
                throw new NotImplementedException("FuncIx=" + functionIndex);
            }
            return result;
        }
        /**
         * @see https://www.openoffice.org/sc/excelfileformat.pdf
         */
        private static Function[] ProduceFunctions()
        {
            Function[] retval = new Function[368];
            retval[0] = new Count(); // COUNT
            retval[FunctionID.IF] = new IfFunc(); //nominally 1
            retval[2] = LogicalFunction.ISNA; // IsNA
            retval[3] = LogicalFunction.ISERROR; // IsERROR
            retval[FunctionID.SUM] = AggregateFunction.SUM; //nominally 4
            retval[5] = AggregateFunction.AVERAGE; // AVERAGE
            retval[6] = AggregateFunction.MIN; // MIN
            retval[7] = AggregateFunction.MAX; // MAX
            retval[8] = new Row(); // ROW
            retval[9] = new Column(); // COLUMN
            retval[10] = new Na(); // NA
            retval[11] = new Npv(); // NPV
            retval[12] = AggregateFunction.STDEV; // STDEV
            retval[13] = NumericFunction.DOLLAR; // DOLLAR
            retval[14] = new Fixed(); // FIXED
            retval[15] = NumericFunction.SIN; // SIN
            retval[16] = NumericFunction.COS; // COS
            retval[17] = NumericFunction.TAN; // TAN
            retval[18] = NumericFunction.ATAN; // ATAN
            retval[19] = new Pi(); // PI
            retval[20] = NumericFunction.SQRT; // SQRT
            retval[21] = NumericFunction.EXP; // EXP
            retval[22] = NumericFunction.LN; // LN
            retval[23] = NumericFunction.LOG10; // LOG10
            retval[24] = NumericFunction.ABS; // ABS
            retval[25] = NumericFunction.INT; // INT
            retval[26] = NumericFunction.SIGN; // SIGN
            retval[27] = NumericFunction.ROUND; // ROUND
            retval[28] = new Lookup(); // LOOKUP
            retval[29] = new NPOI.SS.Formula.Functions.Index(); // INDEX
            retval[30] = new Rept(); // REPT
            retval[31] = TextFunction.MID; // MID
            retval[32] = TextFunction.LEN; // LEN
            retval[33] = new Value(); // VALUE
            retval[34] = new True(); // TRUE
            retval[35] = new False(); // FALSE
            retval[36] = new And(); // AND
            retval[37] = new Or(); // OR
            retval[38] = new Not(); // NOT
            retval[39] = NumericFunction.MOD; // MOD
            retval[40] = new NotImplementedFunction("DCOUNT"); // DCOUNT
            retval[41] = new NotImplementedFunction("DSUM"); // DSUM
            retval[42] = new NotImplementedFunction("DAVERAGE"); // DAVERAGE
            retval[43] = new DStarRunner(DStarRunner.DStarAlgorithmEnum.DMIN); // DMIN
            retval[44] = new NotImplementedFunction("DMAX"); // DMAX
            retval[45] = new NotImplementedFunction("DSTDEV"); // DSTDEV
            retval[46] = AggregateFunction.VAR; // VAR
            retval[47] = new NotImplementedFunction("DVAR"); // DVAR
            retval[48] = TextFunction.TEXT; // TEXT
            retval[49] = new NotImplementedFunction("LINEST"); // LINEST
            retval[50] = new NotImplementedFunction("TREND"); // TREND
            retval[51] = new NotImplementedFunction("LOGEST"); // LOGEST
            retval[52] = new NotImplementedFunction("GROWTH"); // GROWTH
            retval[53] = new NotImplementedFunction("GOTO"); // GOTO
            retval[54] = new NotImplementedFunction("HALT"); // HALT
            retval[56] = FinanceFunction.PV; // PV
            retval[57] = FinanceFunction.FV; // FV
            retval[58] = FinanceFunction.NPER; // NPER
            retval[59] = FinanceFunction.PMT; // PMT
            retval[60] = new Rate(); // RATE
            retval[61] = new Mirr(); // MIRR
            retval[62] = new Irr(); // IRR
            retval[63] = new Rand(); // RAND
            retval[64] = new Match(); // MATCH
            retval[65] = DateFunc.instance; // DATE
            retval[66] = new TimeFunc(); // TIME
            retval[67] = CalendarFieldFunction.DAY; // DAY
            retval[68] = CalendarFieldFunction.MONTH; // MONTH
            retval[69] = CalendarFieldFunction.YEAR; // YEAR
            retval[70] = WeekdayFunc.instance; // WEEKDAY
            retval[71] = CalendarFieldFunction.HOUR;
            retval[72] = CalendarFieldFunction.MINUTE;
            retval[73] = CalendarFieldFunction.SECOND;
            retval[74] = new Now();
            retval[75] = new NotImplementedFunction("AREAS"); // AREAS
            retval[76] = new Rows(); // ROWS
            retval[77] = new Columns(); // COLUMNS
            retval[FunctionID.OFFSET] = new Offset();  //nominally 78
            retval[79] = new NotImplementedFunction("ABSREF"); // ABSREF
            retval[80] = new NotImplementedFunction("RELREF"); // RELREF
            retval[81] = new NotImplementedFunction("ARGUMENT"); // ARGUMENT
            retval[82] = TextFunction.SEARCH;
            retval[83] = new NotImplementedFunction("TRANSPOSE"); // TRANSPOSE
            retval[84] = new NotImplementedFunction("ERROR"); // ERROR
            retval[85] = new NotImplementedFunction("STEP"); // STEP
            retval[86] = new NotImplementedFunction("TYPE"); // TYPE
            retval[87] = new NotImplementedFunction("ECHO"); // ECHO
            retval[88] = new NotImplementedFunction("SetNAME"); // SetNAME
            retval[89] = new NotImplementedFunction("CALLER"); // CALLER
            retval[90] = new NotImplementedFunction("DEREF"); // DEREF
            retval[91] = new NotImplementedFunction("WINDOWS"); // WINDOWS
            retval[92] = new NotImplementedFunction("SERIES"); // SERIES
            retval[93] = new NotImplementedFunction("DOCUMENTS"); // DOCUMENTS
            retval[94] = new NotImplementedFunction("ACTIVECELL"); // ACTIVECELL
            retval[95] = new NotImplementedFunction("SELECTION"); // SELECTION
            retval[96] = new NotImplementedFunction("RESULT"); // RESULT
            retval[97] = NumericFunction.ATAN2; // ATAN2
            retval[98] = NumericFunction.ASIN; // ASIN
            retval[99] = NumericFunction.ACOS; // ACOS
            retval[FunctionID.CHOOSE] = new Choose(); //nominally 100
            retval[101] = new Hlookup(); // HLOOKUP
            retval[102] = new Vlookup(); // VLOOKUP
            retval[103] = new NotImplementedFunction("LINKS"); // LINKS
            retval[104] = new NotImplementedFunction("INPUT"); // INPUT
            retval[105] = LogicalFunction.ISREF; // IsREF
            retval[106] = new NotImplementedFunction("GetFORMULA"); // GetFORMULA
            retval[107] = new NotImplementedFunction("GetNAME"); // GetNAME
            retval[108] = new NotImplementedFunction("SetVALUE"); // SetVALUE
            retval[109] = NumericFunction.LOG; // LOG
            retval[110] = new NotImplementedFunction("EXEC"); // EXEC
            retval[111] = TextFunction.CHAR; // CHAR
            retval[112] = TextFunction.LOWER; // LOWER
            retval[113] = TextFunction.UPPER; // UPPER
            retval[114] = TextFunction.PROPER; // PROPER
            retval[115] = TextFunction.LEFT; // LEFT
            retval[116] = TextFunction.RIGHT; // RIGHT
            retval[117] = TextFunction.EXACT; // EXACT
            retval[118] = TextFunction.TRIM; // TRIM
            retval[119] = new Replace(); // Replace
            retval[120] = new Substitute(); // SUBSTITUTE
            retval[121] = new Code(); // CODE
            retval[122] = new NotImplementedFunction("NAMES"); // NAMES
            retval[123] = new NotImplementedFunction("DIRECTORY"); // DIRECTORY
            retval[124] = TextFunction.FIND; // Find
            retval[125] = new NotImplementedFunction("CELL"); // CELL
            retval[126] = LogicalFunction.ISERR; // IsERR
            retval[127] = LogicalFunction.ISTEXT; // IsTEXT
            retval[128] = LogicalFunction.ISNUMBER; // IsNUMBER
            retval[129] = LogicalFunction.ISBLANK; // IsBLANK
            retval[130] = new T(); // T
            retval[131] = new NotImplementedFunction("N"); // N
            retval[132] = new NotImplementedFunction("FOPEN"); // FOPEN
            retval[133] = new NotImplementedFunction("FCLOSE"); // FCLOSE
            retval[134] = new NotImplementedFunction("FSIZE"); // FSIZE
            retval[135] = new NotImplementedFunction("FReadLN"); // FReadLN
            retval[136] = new NotImplementedFunction("FRead"); // FRead
            retval[137] = new NotImplementedFunction("FWriteLN"); // FWriteLN
            retval[138] = new NotImplementedFunction("FWrite"); // FWrite
            retval[139] = new NotImplementedFunction("FPOS"); // FPOS
            retval[140] = new NotImplementedFunction("DATEVALUE"); // DATEVALUE
            retval[141] = new NotImplementedFunction("TIMEVALUE"); // TIMEVALUE
            retval[142] = new NotImplementedFunction("SLN"); // SLN
            retval[143] = new NotImplementedFunction("SYD"); // SYD
            retval[144] = new NotImplementedFunction("DDB"); // DDB
            retval[145] = new NotImplementedFunction("GetDEF"); // GetDEF
            retval[146] = new NotImplementedFunction("REFTEXT"); // REFTEXT
            retval[147] = new NotImplementedFunction("TEXTREF"); // TEXTREF
            retval[FunctionID.INDIRECT] = null; // Indirect.Evaluate has different signature
            retval[149] = new NotImplementedFunction("REGISTER"); // REGISTER
            retval[150] = new NotImplementedFunction("CALL"); // CALL
            retval[151] = new NotImplementedFunction("AddBAR"); // AddBAR
            retval[152] = new NotImplementedFunction("AddMENU"); // AddMENU
            retval[153] = new NotImplementedFunction("AddCOMMAND"); // AddCOMMAND
            retval[154] = new NotImplementedFunction("ENABLECOMMAND"); // ENABLECOMMAND
            retval[155] = new NotImplementedFunction("CHECKCOMMAND"); // CHECKCOMMAND
            retval[156] = new NotImplementedFunction("RenameCOMMAND"); // RenameCOMMAND
            retval[157] = new NotImplementedFunction("SHOWBAR"); // SHOWBAR
            retval[158] = new NotImplementedFunction("DELETEMENU"); // DELETEMENU
            retval[159] = new NotImplementedFunction("DELETECOMMAND"); // DELETECOMMAND
            retval[160] = new NotImplementedFunction("GetCHARTITEM"); // GetCHARTITEM
            retval[161] = new NotImplementedFunction("DIALOGBOX"); // DIALOGBOX
            retval[162] = TextFunction.CLEAN; // CLEAN
            retval[163] = new NotImplementedFunction("MDETERM"); // MDETERM
            retval[164] = new NotImplementedFunction("MINVERSE"); // MINVERSE
            retval[165] = new NotImplementedFunction("MMULT"); // MMULT
            retval[166] = new NotImplementedFunction("FILES"); // FILES
            retval[167] = new IPMT();
            retval[168] = new PPMT();
            retval[169] = new Counta(); // COUNTA
            retval[170] = new NotImplementedFunction("CANCELKEY"); // CANCELKEY
            retval[175] = new NotImplementedFunction("INITIATE"); // INITIATE
            retval[176] = new NotImplementedFunction("REQUEST"); // REQUEST
            retval[177] = new NotImplementedFunction("POKE"); // POKE
            retval[178] = new NotImplementedFunction("EXECUTE"); // EXECUTE
            retval[179] = new NotImplementedFunction("TERMINATE"); // TERMINATE
            retval[180] = new NotImplementedFunction("RESTART"); // RESTART
            retval[181] = new NotImplementedFunction("HELP"); // HELP
            retval[182] = new NotImplementedFunction("GetBAR"); // GetBAR
            retval[183] = AggregateFunction.PRODUCT; // PRODUCT
            retval[184] = NumericFunction.FACT; // FACT
            retval[185] = new NotImplementedFunction("GetCELL"); // GetCELL
            retval[186] = new NotImplementedFunction("GetWORKSPACE"); // GetWORKSPACE
            retval[187] = new NotImplementedFunction("GetWINDOW"); // GetWINDOW
            retval[188] = new NotImplementedFunction("GetDOCUMENT"); // GetDOCUMENT
            retval[189] = new NotImplementedFunction("DPRODUCT"); // DPRODUCT
            retval[190] = LogicalFunction.ISNONTEXT; // IsNONTEXT
            retval[191] = new NotImplementedFunction("GetNOTE"); // GetNOTE
            retval[192] = new NotImplementedFunction("NOTE"); // NOTE
            retval[193] = new NotImplementedFunction("STDEVP"); // STDEVP
            retval[194] = AggregateFunction.VARP; // VARP
            retval[195] = new NotImplementedFunction("DSTDEVP"); // DSTDEVP
            retval[196] = new NotImplementedFunction("DVARP"); // DVARP
            retval[197] = NumericFunction.TRUNC; // TRUNC
            retval[198] = LogicalFunction.ISLOGICAL; // IsLOGICAL
            retval[199] = new NotImplementedFunction("DCOUNTA"); // DCOUNTA
            retval[200] = new NotImplementedFunction("DELETEBAR"); // DELETEBAR
            retval[201] = new NotImplementedFunction("UNREGISTER"); // UNREGISTER
            retval[204] = new NotImplementedFunction("USDOLLAR"); // USDOLLAR
            retval[205] = new NotImplementedFunction("FindB"); // FindB
            retval[206] = new NotImplementedFunction("SEARCHB"); // SEARCHB
            retval[207] = new NotImplementedFunction("ReplaceB"); // ReplaceB
            retval[208] = new NotImplementedFunction("LEFTB"); // LEFTB
            retval[209] = new NotImplementedFunction("RIGHTB"); // RIGHTB
            retval[210] = new NotImplementedFunction("MIDB"); // MIDB
            retval[211] = new NotImplementedFunction("LENB"); // LENB
            retval[212] = NumericFunction.ROUNDUP; // ROUNDUP
            retval[213] = NumericFunction.ROUNDDOWN; // ROUNDDOWN
            retval[214] = new NotImplementedFunction("ASC"); // ASC
            retval[215] = new NotImplementedFunction("DBCS"); // DBCS
            retval[216] = new Rank(); // RANK
            retval[219] = new Address(); // AddRESS
            retval[220] = new Days360(); // DAYS360
            retval[221] = new Today(); // TODAY
            retval[222] = new NotImplementedFunction("VDB"); // VDB
            retval[227] = AggregateFunction.MEDIAN; // MEDIAN
            retval[228] = new Sumproduct(); // SUMPRODUCT
            retval[229] = NumericFunction.SINH; // SINH
            retval[230] = NumericFunction.COSH; // COSH
            retval[231] = NumericFunction.TANH; // TANH
            retval[232] = NumericFunction.ASINH; // ASINH
            retval[233] = NumericFunction.ACOSH; // ACOSH
            retval[234] = NumericFunction.ATANH; // ATANH
            retval[235] = new DStarRunner(DStarRunner.DStarAlgorithmEnum.DGET);// DGet
            retval[236] = new NotImplementedFunction("CreateOBJECT"); // CreateOBJECT
            retval[237] = new NotImplementedFunction("VOLATILE"); // VOLATILE
            retval[238] = new NotImplementedFunction("LASTERROR"); // LASTERROR
            retval[239] = new NotImplementedFunction("CUSTOMUNDO"); // CUSTOMUNDO
            retval[240] = new NotImplementedFunction("CUSTOMREPEAT"); // CUSTOMREPEAT
            retval[241] = new NotImplementedFunction("FORMULAConvert"); // FORMULAConvert
            retval[242] = new NotImplementedFunction("GetLINKINFO"); // GetLINKINFO
            retval[243] = new NotImplementedFunction("TEXTBOX"); // TEXTBOX
            retval[244] = new NotImplementedFunction("INFO"); // INFO
            retval[245] = new NotImplementedFunction("GROUP"); // GROUP
            retval[246] = new NotImplementedFunction("GetOBJECT"); // GetOBJECT
            retval[247] = new NotImplementedFunction("DB"); // DB
            retval[248] = new NotImplementedFunction("PAUSE"); // PAUSE
            retval[250] = new NotImplementedFunction("RESUME"); // RESUME
            retval[252] = new NotImplementedFunction("FREQUENCY"); // FREQUENCY
            retval[253] = new NotImplementedFunction("AddTOOLBAR"); // AddTOOLBAR
            retval[254] = new NotImplementedFunction("DELETETOOLBAR"); // DELETETOOLBAR
            retval[FunctionID.EXTERNAL_FUNC] = null; // ExternalFunction is a FreeREfFunction
            retval[256] = new NotImplementedFunction("RESetTOOLBAR"); // RESetTOOLBAR
            retval[257] = new NotImplementedFunction("EVALUATE"); // EVALUATE
            retval[258] = new NotImplementedFunction("GetTOOLBAR"); // GetTOOLBAR
            retval[259] = new NotImplementedFunction("GetTOOL"); // GetTOOL
            retval[260] = new NotImplementedFunction("SPELLINGCHECK"); // SPELLINGCHECK
            retval[261] = new Errortype(); // ERRORTYPE
            retval[262] = new NotImplementedFunction("APPTITLE"); // APPTITLE
            retval[263] = new NotImplementedFunction("WINDOWTITLE"); // WINDOWTITLE
            retval[264] = new NotImplementedFunction("SAVETOOLBAR"); // SAVETOOLBAR
            retval[265] = new NotImplementedFunction("ENABLETOOL"); // ENABLETOOL
            retval[266] = new NotImplementedFunction("PRESSTOOL"); // PRESSTOOL
            retval[267] = new NotImplementedFunction("REGISTERID"); // REGISTERID
            retval[268] = new NotImplementedFunction("GetWORKBOOK"); // GetWORKBOOK
            retval[269] = AggregateFunction.AVEDEV; // AVEDEV
            retval[270] = new NotImplementedFunction("BETADIST"); // BETADIST
            retval[271] = new NotImplementedFunction("GAMMALN"); // GAMMALN
            retval[272] = new NotImplementedFunction("BETAINV"); // BETAINV
            retval[273] = new NotImplementedFunction("BINOMDIST"); // BINOMDIST
            retval[274] = new NotImplementedFunction("CHIDIST"); // CHIDIST
            retval[275] = new NotImplementedFunction("CHIINV"); // CHIINV
            retval[276] = NumericFunction.COMBIN; // COMBIN
            retval[277] = new NotImplementedFunction("CONFIDENCE"); // CONFIDENCE
            retval[278] = new NotImplementedFunction("CRITBINOM"); // CRITBINOM
            retval[279] = new Even(); // EVEN
            retval[280] = new NotImplementedFunction("EXPONDIST"); // EXPONDIST
            retval[281] = new NotImplementedFunction("FDIST"); // FDIST
            retval[282] = new NotImplementedFunction("FINV"); // FINV
            retval[283] = new NotImplementedFunction("FISHER"); // FISHER
            retval[284] = new NotImplementedFunction("FISHERINV"); // FISHERINV
            retval[285] = NumericFunction.FLOOR; // FLOOR
            retval[286] = new NotImplementedFunction("GAMMADIST"); // GAMMADIST
            retval[287] = new NotImplementedFunction("GAMMAINV"); // GAMMAINV
            retval[288] = NumericFunction.CEILING; // CEILING
            retval[289] = new NotImplementedFunction("HYPGEOMDIST"); // HYPGEOMDIST
            retval[290] = new NotImplementedFunction("LOGNORMDIST"); // LOGNORMDIST
            retval[291] = new NotImplementedFunction("LOGINV"); // LOGINV
            retval[292] = new NotImplementedFunction("NEGBINOMDIST"); // NEGBINOMDIST
            retval[293] = new NotImplementedFunction("NORMDIST"); // NORMDIST
            retval[294] = new NotImplementedFunction("NORMSDIST"); // NORMSDIST
            retval[295] = new NotImplementedFunction("NORMINV"); // NORMINV
            retval[296] = new NotImplementedFunction("NORMSINV"); // NORMSINV
            retval[297] = new NotImplementedFunction("STANDARDIZE"); // STANDARDIZE
            retval[298] = new Odd(); // ODD
            retval[299] = new NotImplementedFunction("PERMUT"); // PERMUT
            retval[300] = NumericFunction.POISSON; // POISSON
            retval[301] = new NotImplementedFunction("TDIST"); // TDIST
            retval[302] = new NotImplementedFunction("WEIBULL"); // WEIBULL
            retval[303] = new Sumxmy2(); // SUMXMY2
            retval[304] = new Sumx2my2(); // SUMX2MY2
            retval[305] = new Sumx2py2(); // SUMX2PY2
            retval[306] = new NotImplementedFunction("CHITEST"); // CHITEST
            retval[307] = new NotImplementedFunction("CORREL"); // CORREL
            retval[308] = new NotImplementedFunction("COVAR"); // COVAR
            retval[309] = new NotImplementedFunction("FORECAST"); // FORECAST
            retval[310] = new NotImplementedFunction("FTEST"); // FTEST
            retval[311] = new Intercept(); // INTERCEPT
            retval[312] = new NotImplementedFunction("PEARSON"); // PEARSON
            retval[313] = new NotImplementedFunction("RSQ"); // RSQ
            retval[314] = new NotImplementedFunction("STEYX"); // STEYX
            retval[315] = new Slope(); // SLOPE
            retval[316] = new NotImplementedFunction("TTEST"); // TTEST
            retval[317] = new NotImplementedFunction("PROB"); // PROB
            retval[318] = AggregateFunction.DEVSQ; // DEVSQ
            retval[319] = new NotImplementedFunction("GEOMEAN"); // GEOMEAN
            retval[320] = new NotImplementedFunction("HARMEAN"); // HARMEAN
            retval[321] = AggregateFunction.SUMSQ; // SUMSQ
            retval[322] = new NotImplementedFunction("KURT"); // KURT
            retval[323] = new NotImplementedFunction("SKEW"); // SKEW
            retval[324] = new NotImplementedFunction("ZTEST"); // ZTEST
            retval[325] = AggregateFunction.LARGE; // LARGE
            retval[326] = AggregateFunction.SMALL; // SMALL
            retval[327] = new NotImplementedFunction("QUARTILE"); // QUARTILE
            retval[328] = AggregateFunction.PERCENTILE; // PERCENTILE
            retval[329] = new NotImplementedFunction("PERCENTRANK"); // PERCENTRANK
            retval[330] = new Mode(); // MODE
            retval[331] = new NotImplementedFunction("TRIMMEAN"); // TRIMMEAN
            retval[332] = new NotImplementedFunction("TINV"); // TINV
            retval[334] = new NotImplementedFunction("MOVIECOMMAND"); // MOVIECOMMAND
            retval[335] = new NotImplementedFunction("GetMOVIE"); // GetMOVIE
            retval[336] = TextFunction.CONCATENATE; // CONCATENATE
            retval[337] = NumericFunction.POWER; // POWER
            retval[338] = new NotImplementedFunction("PIVOTAddDATA"); // PIVOTAddDATA
            retval[339] = new NotImplementedFunction("GetPIVOTTABLE"); // GetPIVOTTABLE
            retval[340] = new NotImplementedFunction("GetPIVOTFIELD"); // GetPIVOTFIELD
            retval[341] = new NotImplementedFunction("GetPIVOTITEM"); // GetPIVOTITEM
            retval[342] = NumericFunction.RADIANS;
            ; // RADIANS
            retval[343] = NumericFunction.DEGREES; // DEGREES
            retval[344] = new Subtotal(); // SUBTOTAL
            retval[345] = new Sumif(); // SUMIF
            retval[346] = new Countif(); // COUNTIF
            retval[347] = new Countblank(); // COUNTBLANK
            retval[348] = new NotImplementedFunction("SCENARIOGet"); // SCENARIOGet
            retval[349] = new NotImplementedFunction("OPTIONSLISTSGet"); // OPTIONSLISTSGet
            retval[350] = new NotImplementedFunction("IsPMT"); // IsPMT
            retval[351] = new NotImplementedFunction("DATEDIF"); // DATEDIF
            retval[352] = new NotImplementedFunction("DATESTRING"); // DATESTRING
            retval[353] = new NotImplementedFunction("NUMBERSTRING"); // NUMBERSTRING
            retval[354] = new Roman(); // ROMAN
            retval[355] = new NotImplementedFunction("OPENDIALOG"); // OPENDIALOG
            retval[356] = new NotImplementedFunction("SAVEDIALOG"); // SAVEDIALOG
            retval[357] = new NotImplementedFunction("VIEWGet"); // VIEWGet
            retval[358] = new NotImplementedFunction("GetPIVOTDATA"); // GetPIVOTDATA
            retval[359] = new Hyperlink(); // HYPERLINK
            retval[360] = new NotImplementedFunction("PHONETIC"); // PHONETIC
            retval[361] = new NotImplementedFunction("AVERAGEA"); // AVERAGEA
            retval[362] = MinaMaxa.MAXA; // MAXA
            retval[363] = MinaMaxa.MINA; // MINA
            retval[364] = new NotImplementedFunction("STDEVPA"); // STDEVPA
            retval[365] = new NotImplementedFunction("VARPA"); // VARPA
            retval[366] = new NotImplementedFunction("STDEVA"); // STDEVA
            retval[367] = new NotImplementedFunction("VARA"); // VARA
            return retval;
        }

        /**
     * Register a new function in runtime.
     *
     * @param name  the function name
     * @param func  the functoin to register
     * @throws ArgumentException if the function is unknown or already  registered.
     * @since 3.8 beta6
     */

        public static void RegisterFunction(String name, Function func)
        {
            FunctionMetadata metaData = FunctionMetadataRegistry.GetFunctionByName(name);
            if (metaData == null)
            {
                if (AnalysisToolPak.IsATPFunction(name))
                {
                    throw new ArgumentException(name + " is a function from the Excel Analysis Toolpack. " +
                                                "Use AnalysisToolpack.RegisterFunction(String name, FreeRefFunction func) instead.");
                }
                else
                {
                    throw new ArgumentException("Unknown function: " + name);
                }
            }

            int idx = metaData.Index;
            if (functions[idx] is NotImplementedFunction)
            {
                functions[idx] = func;
            }
            else
            {
                throw new ArgumentException("POI already implememts " + name +
                                            ". You cannot override POI's implementations of Excel functions");
            }
        }

        /**
         * Returns a collection of function names implemented by POI.
         *
         * @return an array of supported functions
         * @since 3.8 beta6
         */

        public static ReadOnlyCollection<String> GetSupportedFunctionNames()
        {
            List<String> lst = new List<String>();
            for (int i = 0; i < functions.Length; i++)
            {
                Function func = functions[i];
                FunctionMetadata metaData = FunctionMetadataRegistry.GetFunctionByIndex(i);
                if (func != null && !(func is NotImplementedFunction))
                {
                    lst.Add(metaData.Name);
                }
            }
            lst.Add("INDIRECT"); // INDIRECT is a special case
            return lst.AsReadOnly(); // Collections.unmodifiableCollection(lst);
        }

        /**
         * Returns an array of function names NOT implemented by POI.
         *
         * @return an array of not supported functions
         * @since 3.8 beta6
         */

        public static ReadOnlyCollection<String> GetNotSupportedFunctionNames()
        {
            List<String> lst = new List<String>();
            for (int i = 0; i < functions.Length; i++)
            {
                Function func = functions[i];
                if (func != null && (func is NotImplementedFunction))
                {
                    FunctionMetadata metaData = FunctionMetadataRegistry.GetFunctionByIndex(i);
                    lst.Add(metaData.Name);
                }
            }
            lst.Remove("INDIRECT"); // INDIRECT is a special case
            return lst.AsReadOnly(); // Collections.unmodifiableCollection(lst);
        }
    }
}