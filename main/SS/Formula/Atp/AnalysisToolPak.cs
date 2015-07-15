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

using System.Collections.Generic;
using System.Collections.ObjectModel;
using NPOI.SS.Formula.Function;

namespace NPOI.SS.Formula.Atp
{
    using System;
    using System.Collections;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Udf;

    public class NotImplemented : FreeRefFunction
    {
        private String _functionName;

        public NotImplemented(String functionName)
        {
            _functionName = functionName;
        }

        public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec)
        {
            throw new NotImplementedFunctionException(_functionName);
        }
    }

    public class AnalysisToolPak : UDFFinder
    {
        public static UDFFinder instance = new AnalysisToolPak();

        private static Hashtable _functionsByName = CreateFunctionsMap();

        private AnalysisToolPak()
        {
            // no instances of this class
        }

        public override FreeRefFunction FindFunction(String name)
        {
            // functions that are available in Excel 2007+ have a prefix _xlfn.
            // if you save such a .xlsx workbook as .xls
            if (name.StartsWith("_xlfn.")) name = name.Substring(6);

            return (FreeRefFunction)_functionsByName[name.ToUpper()];
        }

        private static Hashtable CreateFunctionsMap()
        {
            Hashtable m = new Hashtable(100);

            r(m, "ACCRINT", null);
            r(m, "ACCRINTM", null);
            r(m, "AMORDEGRC", null);
            r(m, "AMORLINC", null);
            r(m, "AVERAGEIF", null);
            r(m, "AVERAGEIFS", null);
            r(m, "BAHTTEXT", null);
            r(m, "BESSELI", null);
            r(m, "BESSELJ", null);
            r(m, "BESSELK", null);
            r(m, "BESSELY", null);
            r(m, "BIN2DEC", Bin2Dec.instance);
            r(m, "BIN2HEX", null);
            r(m, "BIN2OCT", null);
            r(m, "COMPLEX", Complex.Instance);
            r(m, "CONVERT", null);
            r(m, "COUNTIFS", null);
            r(m, "COUPDAYBS", null);
            r(m, "COUPDAYS", null);
            r(m, "COUPDAYSNC", null);
            r(m, "COUPNCD", null);
            r(m, "COUPNUM", null);
            r(m, "COUPPCD", null);
            r(m, "CUBEKPIMEMBER", null);
            r(m, "CUBEMEMBER", null);
            r(m, "CUBEMEMBERPROPERTY", null);
            r(m, "CUBERANKEDMEMBER", null);
            r(m, "CUBESET", null);
            r(m, "CUBESETCOUNT", null);
            r(m, "CUBEVALUE", null);
            r(m, "CUMIPMT", null);
            r(m, "CUMPRINC", null);
            r(m, "DEC2BIN", Dec2Bin.instance);
            r(m, "DEC2HEX", Dec2Hex.instance);
            r(m, "DEC2OCT", null);
            r(m, "DELTA", Delta.instance);
            r(m, "DISC", null);
            r(m, "DOLLARDE", null);
            r(m, "DOLLARFR", null);
            r(m, "DURATION", null);
            r(m, "EDATE", EDate.Instance);
            r(m, "EFFECT", null);
            r(m, "EOMONTH", EOMonth.instance);
            r(m, "ERF", null);
            r(m, "ERFC", null);
            r(m, "FACTDOUBLE", FactDouble.instance);
            r(m, "FVSCHEDULE", null);
            r(m, "GCD", null);
            r(m, "GESTEP", null);
            r(m, "HEX2BIN", null);
            r(m, "HEX2DEC", Hex2Dec.instance);
            r(m, "HEX2OCT", null);
            r(m, "IFERROR", IfError.Instance);
            r(m, "IMABS", null);
            r(m, "IMAGINARY", Imaginary.instance);
            r(m, "IMARGUMENT", null);
            r(m, "IMCONJUGATE", null);
            r(m, "IMCOS", null);
            r(m, "IMDIV", null);
            r(m, "IMEXP", null);
            r(m, "IMLN", null);
            r(m, "IMLOG10", null);
            r(m, "IMLOG2", null);
            r(m, "IMPOWER", null);
            r(m, "IMPRODUCT", null);
            r(m, "IMREAL", ImReal.instance);
            r(m, "IMSIN", null);
            r(m, "IMSQRT", null);
            r(m, "IMSUB", null);
            r(m, "IMSUM", null);
            r(m, "INTRATE", null);
            r(m, "ISEVEN", ParityFunction.IS_EVEN);
            r(m, "ISODD", ParityFunction.IS_ODD);
            r(m, "JIS", null);
            r(m, "LCM", null);
            r(m, "MDURATION", null);
            r(m, "MROUND", MRound.Instance);
            r(m, "MULTINOMIAL", null);
            r(m, "NETWORKDAYS", NetworkdaysFunction.instance);
            r(m, "NOMINAL", null);
            r(m, "OCT2BIN", null);
            r(m, "OCT2DEC", Oct2Dec.instance);
            r(m, "OCT2HEX", null);
            r(m, "ODDFPRICE", null);
            r(m, "ODDFYIELD", null);
            r(m, "ODDLPRICE", null);
            r(m, "ODDLYIELD", null);
            r(m, "PRICE", null);
            r(m, "PRICEDISC", null);
            r(m, "PRICEMAT", null);
            r(m, "QUOTIENT", Quotient.instance);
            r(m, "RANDBETWEEN", RandBetween.Instance);
            r(m, "RECEIVED", null);
            r(m, "RTD", null);
            r(m, "SERIESSUM", null);
            r(m, "SQRTPI", null);
            r(m, "SUMIFS", Sumifs.instance);
            r(m, "TBILLEQ", null);
            r(m, "TBILLPRICE", null);
            r(m, "TBILLYIELD", null);
            r(m, "WEEKNUM", WeekNum.instance);
            r(m, "WORKDAY", WorkdayFunction.instance);
            r(m, "XIRR", null);
            r(m, "XNPV", null);
            r(m, "YEARFRAC", YearFrac.instance);
            r(m, "YIELD", null);
            r(m, "YIELDDISC", null);
            r(m, "YIELDMAT", null);
            r(m, "COUNTIFS", Countifs.instance);
            return m;
        }

        private static void r(Hashtable m, String functionName, FreeRefFunction pFunc)
        {
            FreeRefFunction func = pFunc == null ? new NotImplemented(functionName) : pFunc;
            m[functionName] = func;
        }

        public static bool IsATPFunction(String name)
        {
            //AnalysisToolPak inst = (AnalysisToolPak)instance;
            return AnalysisToolPak._functionsByName.ContainsKey(name);
        }

        /**
         * Returns a collection of ATP function names implemented by POI.
         *
         * @return an array of supported functions
         * @since 3.8 beta6
         */
        public static ReadOnlyCollection<String> GetSupportedFunctionNames()
        {
            AnalysisToolPak inst = (AnalysisToolPak)instance;
            List<String> lst = new List<String>();
            foreach (String name in AnalysisToolPak._functionsByName.Keys)
            {
                FreeRefFunction func = (FreeRefFunction)AnalysisToolPak._functionsByName[(name)];
                if (func != null && !(func is NotImplemented))
                {
                    lst.Add(name);
                }
            }
            return lst.AsReadOnly(); //Collections.unmodifiableCollection(lst);
        }

        /**
         * Returns a collection of ATP function names NOT implemented by POI.
         *
         * @return an array of not supported functions
         * @since 3.8 beta6
         */
        public static ReadOnlyCollection<String> GetNotSupportedFunctionNames()
        {
            AnalysisToolPak inst = (AnalysisToolPak)instance;
            List<String> lst = new List<String>();
            foreach (String name in AnalysisToolPak._functionsByName.Keys)
            {
                FreeRefFunction func = (FreeRefFunction)AnalysisToolPak._functionsByName[(name)];
                if (func != null && (func is NotImplemented))
                {
                    lst.Add(name);
                }
            }
            return lst.AsReadOnly(); //Collections.unmodifiableCollection(lst);
        }

        /**
         * Register a ATP function in runtime.
         *
         * @param name  the function name
         * @param func  the functoin to register
         * @throws ArgumentException if the function is unknown or already  registered.
         * @since 3.8 beta6
         */
        public static void RegisterFunction(String name, FreeRefFunction func)
        {
            AnalysisToolPak inst = (AnalysisToolPak)instance;
            if (!IsATPFunction(name))
            {
                FunctionMetadata metaData = FunctionMetadataRegistry.GetFunctionByName(name);
                if (metaData != null)
                {
                    throw new ArgumentException(name + " is a built-in Excel function. " +
                            "Use FunctoinEval.RegisterFunction(String name, Function func) instead.");
                }
                else
                {
                    throw new ArgumentException(name + " is not a function from the Excel Analysis Toolpack.");
                }
            }
            FreeRefFunction f = inst.FindFunction(name);
            if (f != null && !(f is NotImplemented))
            {
                throw new ArgumentException("POI already implememts " + name +
                        ". You cannot override POI's implementations of Excel functions");
            }
            if (_functionsByName.ContainsKey(name))
                _functionsByName[name] = func;
            else
                _functionsByName.Add(name, func);
        }
    }
}