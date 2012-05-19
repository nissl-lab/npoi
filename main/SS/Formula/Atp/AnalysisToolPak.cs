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

		public NotImplemented(String functionName) {
			_functionName = functionName;
		}

		public ValueEval Evaluate(ValueEval[] args, OperationEvaluationContext ec) {
			throw new NotImplementedException(_functionName);
		}
    };

    public class AnalysisToolPak:UDFFinder
    {
        public static UDFFinder instance = new AnalysisToolPak();

        private static Hashtable _functionsByName = CreateFunctionsMap();

        private AnalysisToolPak()
        {
            // no instances of this class
        }

        public override FreeRefFunction FindFunction(String name)
        {
            return (FreeRefFunction)_functionsByName[name];
        }

        private static Hashtable CreateFunctionsMap()
        {
            Hashtable m = new Hashtable(100);

            r(m, "ACCRINT", null);
            r(m, "ACCRINTM", null);
            r(m, "AMORDEGRC", null);
            r(m, "AMORLINC", null);
            r(m, "BESSELI", null);
            r(m, "BESSELJ", null);
            r(m, "BESSELK", null);
            r(m, "BESSELY", null);
            r(m, "BIN2DEC", null);
            r(m, "BIN2HEX", null);
            r(m, "BIN2OCT", null);
            r(m, "CO MPLEX", null);
            r(m, "CONVERT", null);
            r(m, "COUPDAYBS", null);
            r(m, "COUPDAYS", null);
            r(m, "COUPDAYSNC", null);
            r(m, "COUPNCD", null);
            r(m, "COUPNUM", null);
            r(m, "COUPPCD", null);
            r(m, "CUMIPMT", null);
            r(m, "CUMPRINC", null);
            r(m, "DEC2BIN", null);
            r(m, "DEC2HEX", null);
            r(m, "DEC2OCT", null);
            r(m, "DELTA", null);
            r(m, "DISC", null);
            r(m, "DOLLARDE", null);
            r(m, "DOLLARFR", null);
            r(m, "DURATION", null);
            r(m, "EDATE", EDate.Instance);
            r(m, "EFFECT", null);
            r(m, "EOMONTH", null);
            r(m, "ERF", null);
            r(m, "ERFC", null);
            r(m, "FACTDOUBLE", null);
            r(m, "FVSCHEDULE", null);
            r(m, "GCD", null);
            r(m, "GESTEP", null);
            r(m, "HEX2BIN", null);
            r(m, "HEX2DEC", null);
            r(m, "HEX2OCT", null);
            r(m, "IMABS", null);
            r(m, "IMAGINARY", null);
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
            r(m, "IMREAL", null);
            r(m, "IMSIN", null);
            r(m, "IMSQRT", null);
            r(m, "IMSUB", null);
            r(m, "IMSUM", null);
            r(m, "INTRATE", null);
            r(m, "ISEVEN", ParityFunction.IS_EVEN);
            r(m, "ISODD", ParityFunction.IS_ODD);
            r(m, "LCM", null);
            r(m, "MDURATION", null);
            r(m, "MROUND", MRound.Instance);
            r(m, "MULTINOMIAL", null);
            r(m, "NETWORKDAYS", null);
            r(m, "NOMINAL", null);
            r(m, "OCT2BIN", null);
            r(m, "OCT2DEC", null);
            r(m, "OCT2HEX", null);
            r(m, "ODDFPRICE", null);
            r(m, "ODDFYIELD", null);
            r(m, "ODDLPRICE", null);
            r(m, "ODDLYIELD", null);
            r(m, "PRICE", null);
            r(m, "PRICEDISC", null);
            r(m, "PRICEMAT", null);
            r(m, "QUOTIENT", null);
            r(m, "RANDBETWEEN", RandBetween.Instance);
            r(m, "RECEIVED", null);
            r(m, "SERIESSUM", null);
            r(m, "SQRTPI", null);
            r(m, "TBILLEQ", null);
            r(m, "TBILLPRICE", null);
            r(m, "TBILLYIELD", null);
            r(m, "WEEKNUM", null);
            r(m, "WORKDAY", null);
            r(m, "XIRR", null);
            r(m, "XNPV", null);
            r(m, "YEARFRAC", YearFrac.instance);
            r(m, "YIELD", null);
            r(m, "YIELDDISC", null);
            r(m, "YIELDMAT", null);

            return m;
        }

        private static void r(Hashtable m, String functionName, FreeRefFunction pFunc)
        {
            FreeRefFunction func = pFunc == null ? new NotImplemented(functionName) : pFunc;
            m[functionName] = func;
        }
    }
}