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
namespace NPOI.SS.Formula.Eval
{
    using System;
    using System.Collections;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Function;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     *  
     */
    public abstract class FunctionEval 
    {
        /**
         * Some function IDs that require special treatment
         */
        private class FunctionID
        {		/** 1 */
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

        public abstract Eval Evaluate(Eval[] evals, int srcCellRow, short srcCellCol);


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

        private static Function[] ProduceFunctions()
        {
            Function[] retval = new Function[368];
            retval[0] = new Count(); // COUNT
            retval[FunctionID.IF] = new If(); // IF
            retval[2] = LogicalFunction.ISNA; // IsNA
            retval[3] = LogicalFunction.ISERROR; // IsERROR
            retval[FunctionID.SUM] = AggregateFunction.SUM; // SUM
            retval[5] = AggregateFunction.AVERAGE; // AVERAGE
            retval[6] = AggregateFunction.MIN; // MIN
            retval[7] = AggregateFunction.MAX; // MAX
            retval[8] = new Row(); // ROW
            retval[9] = new Column(); // COLUMN
            retval[10] = new Na(); // NA
            retval[11] = new Npv(); // NPV
            retval[12] = AggregateFunction.STDEV; // STDEV
            retval[13] = NumericFunction.DOLLAR; // DOLLAR
            retval[14] = new NotImplementedFunction(); // FIXED
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
            retval[29] = new Index(); // INDEX
            retval[30] = new NotImplementedFunction(); // REPT
            retval[31] = TextFunction.MID; // MID
            retval[32] = TextFunction.LEN; // LEN
            retval[33] = new Value(); // VALUE
            retval[34] = new True(); // TRUE
            retval[35] = new False(); // FALSE
            retval[36] = new And(); // AND
            retval[37] = new Or(); // OR
            retval[38] = new Not(); // NOT
            retval[39] = NumericFunction.MOD; // MOD
            retval[40] = new NotImplementedFunction(); // DCOUNT
            retval[41] = new NotImplementedFunction(); // DSUM
            retval[42] = new NotImplementedFunction(); // DAVERAGE
            retval[43] = new NotImplementedFunction(); // DMIN
            retval[44] = new NotImplementedFunction(); // DMAX
            retval[45] = new NotImplementedFunction(); // DSTDEV
            retval[46] = new NotImplementedFunction(); // VAR
            retval[47] = new NotImplementedFunction(); // DVAR
            retval[48] = TextFunction.TEXT; // TEXT
            retval[49] = new NotImplementedFunction(); // LINEST
            retval[50] = new NotImplementedFunction(); // TREND
            retval[51] = new NotImplementedFunction(); // LOGEST
            retval[52] = new NotImplementedFunction(); // GROWTH
            retval[53] = new NotImplementedFunction(); // GOTO
            retval[54] = new NotImplementedFunction(); // HALT
            retval[56] = new NotImplementedFunction(); // PV
            retval[57] = FinanceFunction.FV; // FV
            retval[58] = FinanceFunction.NPER; // NPER
            retval[59] = FinanceFunction.PMT; // PMT
            retval[60] = new NotImplementedFunction(); // RATE
            retval[61] = new NotImplementedFunction(); // MIRR
            retval[62] = new NotImplementedFunction(); // IRR
            retval[63] = new Rand(); // RAND
            retval[64] = new Match(); // MATCH
            retval[65] = DateFunc.instance; // DATE
            retval[66] = new TimeFunc(); // TIME
            retval[67] = CalendarFieldFunction.DAY; // DAY
            retval[68] = CalendarFieldFunction.MONTH; // MONTH
            retval[69] = CalendarFieldFunction.YEAR; // YEAR
            retval[70] = new NotImplementedFunction(); // WEEKDAY
            retval[71] = CalendarFieldFunction.HOUR;
            retval[72] = CalendarFieldFunction.MINUTE;
            retval[73] = CalendarFieldFunction.SECOND;
            retval[74] = new Now();
            retval[75] = new NotImplementedFunction(); // AREAS
            retval[76] = new Rows(); // ROWS
            retval[77] = new Columns(); // COLUMNS
            retval[FunctionID.OFFSET] = new Offset(); // Offset.Evaluate has a different signature
            retval[79] = new NotImplementedFunction(); // ABSREF
            retval[80] = new NotImplementedFunction(); // RELREF
            retval[81] = new NotImplementedFunction(); // ARGUMENT
            retval[82] = TextFunction.SEARCH;
            retval[83] = new NotImplementedFunction(); // TRANSPOSE
            retval[84] = new NotImplementedFunction(); // ERROR
            retval[85] = new NotImplementedFunction(); // STEP
            retval[86] = new NotImplementedFunction(); // TYPE
            retval[87] = new NotImplementedFunction(); // ECHO
            retval[88] = new NotImplementedFunction(); // SetNAME
            retval[89] = new NotImplementedFunction(); // CALLER
            retval[90] = new NotImplementedFunction(); // DEREF
            retval[91] = new NotImplementedFunction(); // WINDOWS
            retval[92] = new NotImplementedFunction(); // SERIES
            retval[93] = new NotImplementedFunction(); // DOCUMENTS
            retval[94] = new NotImplementedFunction(); // ACTIVECELL
            retval[95] = new NotImplementedFunction(); // SELECTION
            retval[96] = new NotImplementedFunction(); // RESULT
            retval[97] = NumericFunction.ATAN2; // ATAN2
            retval[98] = NumericFunction.ASIN; // ASIN
            retval[99] = NumericFunction.ACOS; // ACOS
            retval[FunctionID.CHOOSE] = new Choose();
            retval[101] = new Hlookup(); // HLOOKUP
            retval[102] = new Vlookup(); // VLOOKUP
            retval[103] = new NotImplementedFunction(); // LINKS
            retval[104] = new NotImplementedFunction(); // INPUT
            retval[105] = LogicalFunction.ISREF; // IsREF
            retval[106] = new NotImplementedFunction(); // GetFORMULA
            retval[107] = new NotImplementedFunction(); // GetNAME
            retval[108] = new NotImplementedFunction(); // SetVALUE
            retval[109] = NumericFunction.LOG; // LOG
            retval[110] = new NotImplementedFunction(); // EXEC
            retval[111] = TextFunction.CHAR; // CHAR
            retval[112] = TextFunction.LOWER; // LOWER
            retval[113] = TextFunction.UPPER; // UPPER
            retval[114] = new NotImplementedFunction(); // PROPER
            retval[115] = TextFunction.LEFT; // LEFT
            retval[116] = TextFunction.RIGHT; // RIGHT
            retval[117] = TextFunction.EXACT; // EXACT
            retval[118] = TextFunction.TRIM; // TRIM
            retval[119] = new Replace(); // Replace
            retval[120] = new Substitute(); // SUBSTITUTE
            retval[121] = new NotImplementedFunction(); // CODE
            retval[122] = new NotImplementedFunction(); // NAMES
            retval[123] = new NotImplementedFunction(); // DIRECTORY
            retval[124] = TextFunction.FIND; // Find
            retval[125] = new NotImplementedFunction(); // CELL
            retval[126] = new NotImplementedFunction(); // IsERR
            retval[127] = LogicalFunction.ISTEXT; // IsTEXT
            retval[128] = LogicalFunction.ISNUMBER; // IsNUMBER
            retval[129] = LogicalFunction.ISBLANK; // IsBLANK
            retval[130] = new T(); // T
            retval[131] = new NotImplementedFunction(); // N
            retval[132] = new NotImplementedFunction(); // FOPEN
            retval[133] = new NotImplementedFunction(); // FCLOSE
            retval[134] = new NotImplementedFunction(); // FSIZE
            retval[135] = new NotImplementedFunction(); // FReadLN
            retval[136] = new NotImplementedFunction(); // FRead
            retval[137] = new NotImplementedFunction(); // FWriteLN
            retval[138] = new NotImplementedFunction(); // FWrite
            retval[139] = new NotImplementedFunction(); // FPOS
            retval[140] = new NotImplementedFunction(); // DATEVALUE
            retval[141] = new NotImplementedFunction(); // TIMEVALUE
            retval[142] = new NotImplementedFunction(); // SLN
            retval[143] = new NotImplementedFunction(); // SYD
            retval[144] = new NotImplementedFunction(); // DDB
            retval[145] = new NotImplementedFunction(); // GetDEF
            retval[146] = new NotImplementedFunction(); // REFTEXT
            retval[147] = new NotImplementedFunction(); // TEXTREF
            retval[FunctionID.INDIRECT] = null; // Indirect.Evaluate has different signature
            retval[149] = new NotImplementedFunction(); // REGISTER
            retval[150] = new NotImplementedFunction(); // CALL
            retval[151] = new NotImplementedFunction(); // AddBAR
            retval[152] = new NotImplementedFunction(); // AddMENU
            retval[153] = new NotImplementedFunction(); // AddCOMMAND
            retval[154] = new NotImplementedFunction(); // ENABLECOMMAND
            retval[155] = new NotImplementedFunction(); // CHECKCOMMAND
            retval[156] = new NotImplementedFunction(); // RenameCOMMAND
            retval[157] = new NotImplementedFunction(); // SHOWBAR
            retval[158] = new NotImplementedFunction(); // DELETEMENU
            retval[159] = new NotImplementedFunction(); // DELETECOMMAND
            retval[160] = new NotImplementedFunction(); // GetCHARTITEM
            retval[161] = new NotImplementedFunction(); // DIALOGBOX
            retval[162] = TextFunction.CLEAN; // CLEAN
            retval[163] = new NotImplementedFunction(); // MDETERM
            retval[164] = new NotImplementedFunction(); // MINVERSE
            retval[165] = new NotImplementedFunction(); // MMULT
            retval[166] = new NotImplementedFunction(); // FILES
            retval[167] = new NotImplementedFunction(); // IPMT
            retval[168] = new NotImplementedFunction(); // PPMT
            retval[169] = new Counta(); // COUNTA
            retval[170] = new NotImplementedFunction(); // CANCELKEY
            retval[175] = new NotImplementedFunction(); // INITIATE
            retval[176] = new NotImplementedFunction(); // REQUEST
            retval[177] = new NotImplementedFunction(); // POKE
            retval[178] = new NotImplementedFunction(); // EXECUTE
            retval[179] = new NotImplementedFunction(); // TERMINATE
            retval[180] = new NotImplementedFunction(); // RESTART
            retval[181] = new NotImplementedFunction(); // HELP
            retval[182] = new NotImplementedFunction(); // GetBAR
            retval[183] = AggregateFunction.PRODUCT; // PRODUCT
            retval[184] = NumericFunction.FACT; // FACT
            retval[185] = new NotImplementedFunction(); // GetCELL
            retval[186] = new NotImplementedFunction(); // GetWORKSPACE
            retval[187] = new NotImplementedFunction(); // GetWINDOW
            retval[188] = new NotImplementedFunction(); // GetDOCUMENT
            retval[189] = new NotImplementedFunction(); // DPRODUCT
            retval[190] = LogicalFunction.ISNONTEXT; // IsNONTEXT
            retval[191] = new NotImplementedFunction(); // GetNOTE
            retval[192] = new NotImplementedFunction(); // NOTE
            retval[193] = new NotImplementedFunction(); // STDEVP
            retval[194] = new NotImplementedFunction(); // VARP
            retval[195] = new NotImplementedFunction(); // DSTDEVP
            retval[196] = new NotImplementedFunction(); // DVARP
            retval[197] = NumericFunction.TRUNC; // TRUNC
            retval[198] = LogicalFunction.ISLOGICAL; // IsLOGICAL
            retval[199] = new NotImplementedFunction(); // DCOUNTA
            retval[200] = new NotImplementedFunction(); // DELETEBAR
            retval[201] = new NotImplementedFunction(); // UNREGISTER
            retval[204] = new NotImplementedFunction(); // USDOLLAR
            retval[205] = new NotImplementedFunction(); // FindB
            retval[206] = new NotImplementedFunction(); // SEARCHB
            retval[207] = new NotImplementedFunction(); // ReplaceB
            retval[208] = new NotImplementedFunction(); // LEFTB
            retval[209] = new NotImplementedFunction(); // RIGHTB
            retval[210] = new NotImplementedFunction(); // MIDB
            retval[211] = new NotImplementedFunction(); // LENB
            retval[212] = NumericFunction.ROUNDUP; // ROUNDUP
            retval[213] = NumericFunction.ROUNDDOWN; // ROUNDDOWN
            retval[214] = new NotImplementedFunction(); // ASC
            retval[215] = new NotImplementedFunction(); // DBCS
            retval[216] = new NotImplementedFunction(); // RANK
            retval[219] = new Address(); // AddRESS
            retval[220] = new Days360(); // DAYS360
            retval[221] = new NotImplementedFunction(); // TODAY
            retval[222] = new NotImplementedFunction(); // VDB
            retval[227] = AggregateFunction.MEDIAN; // MEDIAN
            retval[228] = new Sumproduct(); // SUMPRODUCT
            retval[229] = NumericFunction.SINH; // SINH
            retval[230] = NumericFunction.COSH; // COSH
            retval[231] = NumericFunction.TANH; // TANH
            retval[232] = NumericFunction.ASINH; // ASINH
            retval[233] = NumericFunction.ACOSH; // ACOSH
            retval[234] = NumericFunction.ATANH; // ATANH
            retval[235] = new NotImplementedFunction(); // DGet
            retval[236] = new NotImplementedFunction(); // CreateOBJECT
            retval[237] = new NotImplementedFunction(); // VOLATILE
            retval[238] = new NotImplementedFunction(); // LASTERROR
            retval[239] = new NotImplementedFunction(); // CUSTOMUNDO
            retval[240] = new NotImplementedFunction(); // CUSTOMREPEAT
            retval[241] = new NotImplementedFunction(); // FORMULAConvert
            retval[242] = new NotImplementedFunction(); // GetLINKINFO
            retval[243] = new NotImplementedFunction(); // TEXTBOX
            retval[244] = new NotImplementedFunction(); // INFO
            retval[245] = new NotImplementedFunction(); // GROUP
            retval[246] = new NotImplementedFunction(); // GetOBJECT
            retval[247] = new NotImplementedFunction(); // DB
            retval[248] = new NotImplementedFunction(); // PAUSE
            retval[250] = new NotImplementedFunction(); // RESUME
            retval[252] = new NotImplementedFunction(); // FREQUENCY
            retval[253] = new NotImplementedFunction(); // AddTOOLBAR
            retval[254] = new NotImplementedFunction(); // DELETETOOLBAR
            retval[FunctionID.EXTERNAL_FUNC] = null; // ExternalFunction is a FreeREfFunction
            retval[256] = new NotImplementedFunction(); // RESetTOOLBAR
            retval[257] = new NotImplementedFunction(); // EVALUATE
            retval[258] = new NotImplementedFunction(); // GetTOOLBAR
            retval[259] = new NotImplementedFunction(); // GetTOOL
            retval[260] = new NotImplementedFunction(); // SPELLINGCHECK
            retval[261] = new NotImplementedFunction(); // ERRORTYPE
            retval[262] = new NotImplementedFunction(); // APPTITLE
            retval[263] = new NotImplementedFunction(); // WINDOWTITLE
            retval[264] = new NotImplementedFunction(); // SAVETOOLBAR
            retval[265] = new NotImplementedFunction(); // ENABLETOOL
            retval[266] = new NotImplementedFunction(); // PRESSTOOL
            retval[267] = new NotImplementedFunction(); // REGISTERID
            retval[268] = new NotImplementedFunction(); // GetWORKBOOK
            retval[269] = AggregateFunction.AVEDEV; // AVEDEV
            retval[270] = new NotImplementedFunction(); // BETADIST
            retval[271] = new NotImplementedFunction(); // GAMMALN
            retval[272] = new NotImplementedFunction(); // BETAINV
            retval[273] = new NotImplementedFunction(); // BINOMDIST
            retval[274] = new NotImplementedFunction(); // CHIDIST
            retval[275] = new NotImplementedFunction(); // CHIINV
            retval[276] = NumericFunction.COMBIN; // COMBIN
            retval[277] = new NotImplementedFunction(); // CONFIDENCE
            retval[278] = new NotImplementedFunction(); // CRITBINOM
            retval[279] = new Even(); // EVEN
            retval[280] = new NotImplementedFunction(); // EXPONDIST
            retval[281] = new NotImplementedFunction(); // FDIST
            retval[282] = new NotImplementedFunction(); // FINV
            retval[283] = new NotImplementedFunction(); // FISHER
            retval[284] = new NotImplementedFunction(); // FISHERINV
            retval[285] = NumericFunction.FLOOR; // FLOOR
            retval[286] = new NotImplementedFunction(); // GAMMADIST
            retval[287] = new NotImplementedFunction(); // GAMMAINV
            retval[288] = NumericFunction.CEILING; // CEILING
            retval[289] = new NotImplementedFunction(); // HYPGEOMDIST
            retval[290] = new NotImplementedFunction(); // LOGNORMDIST
            retval[291] = new NotImplementedFunction(); // LOGINV
            retval[292] = new NotImplementedFunction(); // NEGBINOMDIST
            retval[293] = new NotImplementedFunction(); // NORMDIST
            retval[294] = new NotImplementedFunction(); // NORMSDIST
            retval[295] = new NotImplementedFunction(); // NORMINV
            retval[296] = new NotImplementedFunction(); // NORMSINV
            retval[297] = new NotImplementedFunction(); // STANDARDIZE
            retval[298] = new Odd(); // ODD
            retval[299] = new NotImplementedFunction(); // PERMUT
            retval[300] = new NotImplementedFunction(); // POISSON
            retval[301] = new NotImplementedFunction(); // TDIST
            retval[302] = new NotImplementedFunction(); // WEIBULL
            retval[303] = new Sumxmy2(); // SUMXMY2
            retval[304] = new Sumx2my2(); // SUMX2MY2
            retval[305] = new Sumx2py2(); // SUMX2PY2
            retval[306] = new NotImplementedFunction(); // CHITEST
            retval[307] = new NotImplementedFunction(); // CORREL
            retval[308] = new NotImplementedFunction(); // COVAR
            retval[309] = new NotImplementedFunction(); // FORECAST
            retval[310] = new NotImplementedFunction(); // FTEST
            retval[311] = new NotImplementedFunction(); // INTERCEPT
            retval[312] = new NotImplementedFunction(); // PEARSON
            retval[313] = new NotImplementedFunction(); // RSQ
            retval[314] = new NotImplementedFunction(); // STEYX
            retval[315] = new NotImplementedFunction(); // SLOPE
            retval[316] = new NotImplementedFunction(); // TTEST
            retval[317] = new NotImplementedFunction(); // PROB
            retval[318] = AggregateFunction.DEVSQ; // DEVSQ
            retval[319] = new NotImplementedFunction(); // GEOMEAN
            retval[320] = new NotImplementedFunction(); // HARMEAN
            retval[321] = AggregateFunction.SUMSQ; // SUMSQ
            retval[322] = new NotImplementedFunction(); // KURT
            retval[323] = new NotImplementedFunction(); // SKEW
            retval[324] = new NotImplementedFunction(); // ZTEST
            retval[325] = AggregateFunction.LARGE; // LARGE
            retval[326] = AggregateFunction.SMALL; // SMALL
            retval[327] = new NotImplementedFunction(); // QUARTILE
            retval[328] = new NotImplementedFunction(); // PERCENTILE
            retval[329] = new NotImplementedFunction(); // PERCENTRANK
            retval[330] = new Mode(); // MODE
            retval[331] = new NotImplementedFunction(); // TRIMMEAN
            retval[332] = new NotImplementedFunction(); // TINV
            retval[334] = new NotImplementedFunction(); // MOVIECOMMAND
            retval[335] = new NotImplementedFunction(); // GetMOVIE
            retval[336] = TextFunction.CONCATENATE; // CONCATENATE
            retval[337] = NumericFunction.POWER; // POWER
            retval[338] = new NotImplementedFunction(); // PIVOTAddDATA
            retval[339] = new NotImplementedFunction(); // GetPIVOTTABLE
            retval[340] = new NotImplementedFunction(); // GetPIVOTFIELD
            retval[341] = new NotImplementedFunction(); // GetPIVOTITEM
            retval[342] = NumericFunction.RADIANS; ; // RADIANS
            retval[343] = NumericFunction.DEGREES; // DEGREES
            retval[344] = new Subtotal(); // SUBTOTAL
            retval[345] = new Sumif(); // SUMIF
            retval[346] = new Countif(); // COUNTIF
            retval[347] = new Countblank(); // COUNTBLANK
            retval[348] = new NotImplementedFunction(); // SCENARIOGet
            retval[349] = new NotImplementedFunction(); // OPTIONSLISTSGet
            retval[350] = new NotImplementedFunction(); // IsPMT
            retval[351] = new NotImplementedFunction(); // DATEDIF
            retval[352] = new NotImplementedFunction(); // DATESTRING
            retval[353] = new NotImplementedFunction(); // NUMBERSTRING
            retval[354] = new NotImplementedFunction(); // ROMAN
            retval[355] = new NotImplementedFunction(); // OPENDIALOG
            retval[356] = new NotImplementedFunction(); // SAVEDIALOG
            retval[357] = new NotImplementedFunction(); // VIEWGet
            retval[358] = new NotImplementedFunction(); // GetPIVOTDATA
            retval[359] = new NotImplementedFunction(); // HYPERLINK
            retval[360] = new NotImplementedFunction(); // PHONETIC
            retval[361] = new NotImplementedFunction(); // AVERAGEA
            retval[362] = new Maxa(); // MAXA
            retval[363] = new Mina(); // MINA
            retval[364] = new NotImplementedFunction(); // STDEVPA
            retval[365] = new NotImplementedFunction(); // VARPA
            retval[366] = new NotImplementedFunction(); // STDEVA
            retval[367] = new NotImplementedFunction(); // VARA
            return retval;
        }
    }
}