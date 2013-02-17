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

namespace NPOI.SS.Formula.Function
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Text.RegularExpressions;
    using NPOI.SS.Formula.PTG;
    using System.Globalization;

    /**
     * Converts the text meta-data file into a <c>FunctionMetadataRegistry</c>
     * 
     * @author Josh Micich
     */
    class FunctionMetadataReader
    {

        private const String METADATA_FILE_NAME = "functionMetadata.txt";

        /** plain ASCII text metadata file uses three dots for ellipsis */
        private const string ELLIPSIS = "...";

        private const string TAB_DELIM_PATTERN = @"\t";
        private const string SPACE_DELIM_PATTERN = @"\s";
        private static readonly byte[] EMPTY_BYTE_ARRAY = { };

        private static readonly string[] DIGIT_ENDING_FUNCTION_NAMES = {
		// Digits at the end of a function might be due to a left-over footnote marker.
		// except in these cases
		"LOG10", "ATAN2", "DAYS360", "SUMXMY2", "SUMX2MY2", "SUMX2PY2",
	};
		private static List<string> DIGIT_ENDING_FUNCTION_NAMES_Set = new List<string> (DIGIT_ENDING_FUNCTION_NAMES);

        public static FunctionMetadataRegistry CreateRegistry()
        {
            using (StreamReader br = new StreamReader (typeof (FunctionMetadataReader).Assembly.GetManifestResourceStream (METADATA_FILE_NAME)))
            {

                FunctionDataBuilder fdb = new FunctionDataBuilder(400);

                try
                {
                    while (true)
                    {
                        String line = br.ReadLine();
                        if (line == null)
                        {
                            break;
                        }
                        if (line.Length < 1 || line[0] == '#')
                        {
                            continue;
                        }
                        String TrimLine = line.Trim();
                        if (TrimLine.Length < 1)
                        {
                            continue;
                        }
                        ProcessLine(fdb, line);
                    }
                }
                catch (IOException)
                {
                    throw;
                }

                return fdb.Build();
            }
        }

        private static void ProcessLine(FunctionDataBuilder fdb, String line)
        {

            Regex regex = new Regex(TAB_DELIM_PATTERN);
            String[] parts = regex.Split(line);
            if (parts.Length != 8)
            {
                throw new Exception("Bad line format '" + line + "' - expected 8 data fields");
            }
            int functionIndex = ParseInt(parts[0]);
            String functionName = parts[1];
            int minParams = ParseInt(parts[2]);
            int maxParams = ParseInt(parts[3]);
            byte returnClassCode = ParseReturnTypeCode(parts[4]);
            byte[] parameterClassCodes = ParseOperandTypeCodes(parts[5]);
            // 6 IsVolatile
            bool hasNote = parts[7].Length > 0;

            ValidateFunctionName(functionName);
            // TODO - make POI use IsVolatile
            fdb.Add(functionIndex, functionName, minParams, maxParams,
                    returnClassCode, parameterClassCodes, hasNote);
        }


        private static byte ParseReturnTypeCode(String code)
        {
            if (code.Length == 0)
            {
                return Ptg.CLASS_REF; // happens for GetPIVOTDATA
            }
            return ParseOperandTypeCode(code);
        }

        private static byte[] ParseOperandTypeCodes(String codes)
        {
            if (codes.Length < 1)
            {
                return EMPTY_BYTE_ARRAY; // happens for GetPIVOTDATA
            }
            if (IsDash(codes))
            {
                // '-' means empty:
                return EMPTY_BYTE_ARRAY;
            }
            Regex regex = new Regex(SPACE_DELIM_PATTERN);
            String[] array = regex.Split(codes);
            int nItems = array.Length;
            if (ELLIPSIS.Equals(array[nItems - 1]))
            {
                // ellipsis is optional, and ignored
                // (all Unspecified params are assumed to be the same as the last)
                nItems--;
            }
            byte[] result = new byte[nItems];
            for (int i = 0; i < nItems; i++)
            {
                result[i] = ParseOperandTypeCode(array[i]);
            }
            return result;
        }

        private static bool IsDash(String codes)
        {
            if (codes.Length == 1)
            {
                switch (codes[0])
                {
                    case '-':
                        return true;
                }
            }
            return false;
        }

        private static byte ParseOperandTypeCode(String code)
        {
            if (code.Length != 1)
            {
                throw new Exception("Bad operand type code format '" + code + "' expected single char");
            }
            switch (code[0])
            {
                case 'V': return Ptg.CLASS_VALUE;
                case 'R': return Ptg.CLASS_REF;
                case 'A': return Ptg.CLASS_ARRAY;
            }
            throw new ArgumentException("Unexpected operand type code '" + code + "' (" + (int)code[0] + ")");
        }

        /**
         * Makes sure that footnote digits from the original OOO document have not been accidentally 
         * left behind
         */
        private static void ValidateFunctionName(String functionName)
        {
            int len = functionName.Length;
            int ix = len - 1;
            if (!Char.IsDigit(functionName[ix]))
            {
                return;
            }
            while (ix >= 0)
            {
                if (!Char.IsDigit(functionName[ix]))
                {
                    break;
                }
                ix--;
            }
            if (DIGIT_ENDING_FUNCTION_NAMES_Set.Contains(functionName))
            {
                return;
            }
            throw new Exception("Invalid function name '" + functionName
                    + "' (is footnote number incorrectly Appended)");
        }

        private static int ParseInt(String valStr)
        {
            return int.Parse(valStr, CultureInfo.InvariantCulture);
        }
    }
}