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

using System;
using NPOI.SS.UserModel;
using NPOI.SS;
using System.Text;

namespace TestCases.SS.Util
{

    /**
     * Contains specific examples of <c>double</c> values and their rendering in Excel.
     *
     * @author Josh Micich
     */
    public sealed class NumberToTextConversionExamples
    {

        private NumberToTextConversionExamples()
        {
            // no instances of this class
        }
        /// <summary>
        /// rewrite according the java class;
        /// </summary>
        public class ExampleConversion
        {
            private String _cSharpRendering;
            private String _excelRendering;
            private double _doubleValue;
            private long _rawDoubleBits;

            public ExampleConversion(long rawDoubleBits, String cSharpRendering, String excelRendering)
            {
                //double d = Double.longBitsToDouble(rawDoubleBits);
                double d = BitConverter.Int64BitsToDouble(rawDoubleBits);
                if ("NaN".Equals(cSharpRendering))
                {
                    if (!Double.IsNaN(d))
                    {
                        throw new ArgumentException("value must be NaN");
                    }
                }
                else if (double.IsInfinity(d))
                {
                    if (!Double.IsInfinity(d))
                    {
                        throw new ArgumentException("value must be Infinity");
                    }
                }
                else
                {
                    if (Double.IsNaN(d))
                    {
                        throw new ArgumentException("value must not be NaN");
                    }

                    // just to be dead sure test conversion in java both ways
                    bool javaToStringOk = cSharpRendering.Equals(d.ToString("R"));
                    bool javaParseOk = double.Parse(cSharpRendering) == d;
                    if (!javaToStringOk || !javaParseOk)
                    {
                        String msgA = "Specified rawDoubleBits " + DoubleToHexString(d) + " encodes to double '" + d.ToString("R") + "'.";
                        String msgB = "Specified cSharpRendering '" + cSharpRendering + "' parses as double with rawDoubleBits "
                            + DoubleToHexString(double.Parse(cSharpRendering));
                        System.Console.WriteLine(msgA);
                        System.Console.WriteLine(msgB);
                        throw new Exception(msgA + msgB);
                    }
                }
                _rawDoubleBits = rawDoubleBits;
                _cSharpRendering = cSharpRendering;
                _excelRendering = excelRendering;
                _doubleValue = d;
            }
            private static String DoubleToHexString(double d)
            {
                return "0x" + BitConverter.DoubleToInt64Bits(d).ToString("X") + "L";
            }
            public String CSharpRendering
            {
                get
                {
                    return _cSharpRendering;
                }
            }
            public String ExcelRendering
            {
                get
                {
                    return _excelRendering;
                }
            }
            public double DoubleValue
            {
                get
                {
                    return _doubleValue;
                }
            }
            public bool IsNaN
            {
                get
                {
                    return Double.IsNaN(_doubleValue);
                }
            }
            public long RawDoubleBits
            {
                get
                {
                    return _rawDoubleBits;
                }
            }
        }


        /**
         * Number rendering examples as observed from Excel.
         * TODO - some are currently disabled because POI does not pass these cases yet
         */
        private static ExampleConversion[] examples = {

		// basic numbers
        Ec(0x0000000000000000L, "0", "0"),//Ec(0x0000000000000000L, "0.0", "0"),
        Ec(0x3FF0000000000000L, "1", "1"),//Ec(0x3FF0000000000000L, "1.0", "1"),
        Ec(0x3FF00068DB8BAC71L, "1.0001", "1.0001"),
        Ec(0x4087A00000000000L, "756", "756"),//Ec(0x4087A00000000000L, "756.0", "756"),
        Ec(0x401E3D70A3D70A3DL, "7.56", "7.56"),

        Ec(0x405EDD3C07FB4C8CL, "123.4567890123455",  "123.456789012345"),
        Ec(0x405EDD3C07FB4C99L, "123.45678901234568", "123.456789012346"),
        Ec(0x405EDD3C07FB4CAEL, "123.45678901234598", "123.456789012346"),
        Ec(0x4132D687E3DF2180L, "1234567.8901234567", "1234567.89012346"),

        Ec(0x3F543A272D9E0E49L, "0.001234567890123455",  "0.00123456789012345"),
        Ec(0x3F543A272D9E0E4AL, "0.0012345678901234552", "0.00123456789012346"),
        Ec(0x3F543A272D9E0E55L, "0.0012345678901234576", "0.00123456789012346"),
        Ec(0x3F543A272D9E0E72L, "0.0012345678901234639", "0.00123456789012346"),
        Ec(0x3F543A272D9E0E76L, "0.0012345678901234647", "0.00123456789012346"),
        Ec(0x3F543A272D9E0E77L, "0.001234567890123465",  "0.00123456789012346"),

        Ec(0x3F543A272D9E0E78L, "0.0012345678901234652", "0.00123456789012347"),

       
        Ec(0x3F543A272D9E0EA5L, "0.0012345678901234749",  "0.00123456789012347"), //java: Ec(0x3F543A272D9E0EA5L, "0.001234567890123475",  "0.00123456789012347"),
        Ec(0x3F543A272D9E0EA6L, "0.0012345678901234751",  "0.00123456789012348"),

        
        Ec(0x544CE6345CF3209CL, "1.2345678901234549E+98",  "1.23456789012345E+98"),
        Ec(0x544CE6345CF3209DL, "1.2345678901234551E+98",   "1.23456789012346E+98"), //Ec(0x544CE6345CF3209DL, "1.234567890123455E98",   "1.23456789012346E+98"),
        Ec(0x544CE6345CF320DEL, "1.2345678901234649E+98",  "1.23456789012346E+98"),
        
        Ec(0x544CE6345CF320DFL, "1.2345678901234651E+98",  "1.23456789012347E+98"),  //Ec(0x544CE6345CF320DFL, "1.2345678901234651E98",  "1.23456789012347E+98"),
        Ec(0x544CE6345CF32120L, "1.2345678901234749E+98",   "1.23456789012347E+98"),//Ec(0x544CE6345CF32120L, "1.234567890123475E+98",   "1.23456789012347E+98"),
        Ec(0x544CE6345CF32121L, "1.2345678901234751E+98",  "1.23456789012348E+98"),


        Ec(0x54820FE0BA17F5E9L, "1.23456789012355E+99",    "1.2345678901236E+99"),
        Ec(0x54820FE0BA17F5EAL, "1.2345678901235502E+99",  "1.2345678901236E+99"),
        Ec(0x54820FE0BA17F784L, "1.2345678901236498E+99",  "1.2345678901237E+99"),
        Ec(0x54820FE0BA17F785L, "1.23456789012365E+99",    "1.2345678901237E+99"),
        Ec(0x54820FE0BA17F920L, "1.2345678901237498E+99",  "1.2345678901238E+99"),
        Ec(0x54820FE0BA17F921L, "1.23456789012375E+99",    "1.2345678901238E+99"),


        // transitions around the E98,E99,E100 boundaries
        Ec(0x547D42AEA2879F19L,"9.9999999999999742E+98",  "9.99999999999997E+98"), //Ec(0x547D42AEA2879F19L,"9.999999999999974E+98",  "9.99999999999997E+98"),
        Ec(0x547D42AEA2879F1AL,"9.9999999999999754E+98",  "9.99999999999998E+98"),//Ec(0x547D42AEA2879F1AL,"9.999999999999975E+98",  "9.99999999999998E+98"),
        Ec(0x547D42AEA2879F21L,"9.9999999999999839E+98",  "9.99999999999998E+98"),//Ec(0x547D42AEA2879F21L,"9.999999999999984E+98",  "9.99999999999998E+98"),
        Ec(0x547D42AEA2879F22L,"9.9999999999999851E+98",  "9.99999999999999E+98"),//Ec(0x547D42AEA2879F22L,"9.999999999999985E+98",  "9.99999999999999E+98"),
        Ec(0x547D42AEA2879F2AL,"9.9999999999999948E+98",  "9.99999999999999E+98"), //Ec(0x547D42AEA2879F2AL,"9.999999999999995E+98",  "9.99999999999999E+98"),
        Ec(0x547D42AEA2879F2BL,"9.999999999999996E+98",  "1E+99"),
        Ec(0x547D42AEA287A0A0L,"1.0000000000000449E+99", "1E+99"),
        Ec(0x547D42AEA287A0A1L,"1.000000000000045E+99",  "1.0000000000001E+99"),
        Ec(0x547D42AEA287A3D8L,"1.0000000000001449E+99", "1.0000000000001E+99"),
        Ec(0x547D42AEA287A3D9L,"1.0000000000001451E+99",  "1.0000000000002E+99"),//Ec(0x547D42AEA287A3D9L,"1.000000000000145E+99",  "1.0000000000002E+99"),
        Ec(0x547D42AEA287A710L,"1.000000000000245E+99",  "1.0000000000002E+99"),
        Ec(0x547D42AEA287A711L,"1.0000000000002451E+99", "1.0000000000003E+99"),


        Ec(0x54B249AD2594C2F9L,"9.9999999999997437E+99",  "9.9999999999997E+99"),//Ec(0x54B249AD2594C2F9L,"9.999999999999744E+99",  "9.9999999999997E+99"),
        Ec(0x54B249AD2594C2FAL,"9.9999999999997457E+99",  "9.9999999999998E+99"),//Ec(0x54B249AD2594C2FAL,"9.999999999999746E+99",  "9.9999999999998E+99"),
        Ec(0x54B249AD2594C32DL,"9.9999999999998447E+99",  "9.9999999999998E+99"),//Ec(0x54B249AD2594C32DL,"9.999999999999845E+99",  "9.9999999999998E+99"),
        Ec(0x54B249AD2594C32EL,"9.9999999999998467E+99",  "9.9999999999999E+99"),//Ec(0x54B249AD2594C32EL,"9.999999999999847E+99",  "9.9999999999999E+99"),
        Ec(0x54B249AD2594C360L,"9.9999999999999438E+99",  "9.9999999999999E+99"),//Ec(0x54B249AD2594C360L,"9.999999999999944E+99",  "9.9999999999999E+99"),
        Ec(0x54B249AD2594C361L,"9.9999999999999458E+99",  "1E+100"),//Ec(0x54B249AD2594C361L,"9.999999999999946E+99",  "1E+100"),
        Ec(0x54B249AD2594C464L,"1.0000000000000449E+100","1E+100"),
        Ec(0x54B249AD2594C465L,"1.0000000000000451E+100", "1.0000000000001E+100"),//Ec(0x54B249AD2594C465L,"1.000000000000045E+100", "1.0000000000001E+100"),
        Ec(0x54B249AD2594C667L,"1.0000000000001449E+100", "1.0000000000001E+100"),//Ec(0x54B249AD2594C667L,"1.000000000000145E+100", "1.0000000000001E+100"),
        Ec(0x54B249AD2594C668L,"1.0000000000001451E+100","1.0000000000002E+100"),
        Ec(0x54B249AD2594C86AL,"1.000000000000245E+100", "1.0000000000002E+100"),
        Ec(0x54B249AD2594C86BL,"1.0000000000002452E+100","1.0000000000003E+100"),


        Ec(0x2B95DF5CA28EF4A8L,"1.0000000000000251E-98","1.00000000000003E-98"),
        Ec(0x2B95DF5CA28EF4A7L,"1.000000000000025E-98", "1.00000000000002E-98"),
        Ec(0x2B95DF5CA28EF46AL,"1.000000000000015E-98", "1.00000000000002E-98"),
        Ec(0x2B95DF5CA28EF469L,"1.0000000000000149E-98","1.00000000000001E-98"),
        Ec(0x2B95DF5CA28EF42DL,"1.0000000000000051E-98","1.00000000000001E-98"),
        Ec(0x2B95DF5CA28EF42CL,"1.000000000000005E-98", "1E-98"),
        Ec(0x2B95DF5CA28EF3ECL,"9.9999999999999458E-99", "1E-98"),//Ec(0x2B95DF5CA28EF3ECL,"9.999999999999946E-99", "1E-98"),
        Ec(0x2B95DF5CA28EF3EBL,"9.9999999999999442E-99", "9.9999999999999E-99"),//Ec(0x2B95DF5CA28EF3EBL,"9.999999999999944E-99", "9.9999999999999E-99"),
        Ec(0x2B95DF5CA28EF3AEL,"9.9999999999998451E-99", "9.9999999999999E-99"),//Ec(0x2B95DF5CA28EF3AEL,"9.999999999999845E-99", "9.9999999999999E-99"),
        Ec(0x2B95DF5CA28EF3ADL,"9.9999999999998435E-99", "9.9999999999998E-99"),//Ec(0x2B95DF5CA28EF3ADL,"9.999999999999843E-99", "9.9999999999998E-99"),
        Ec(0x2B95DF5CA28EF371L,"9.999999999999746E-99", "9.9999999999998E-99"),
        Ec(0x2B95DF5CA28EF370L,"9.9999999999997444E-99", "9.9999999999997E-99"),//Ec(0x2B95DF5CA28EF370L,"9.999999999999744E-99", "9.9999999999997E-99"),


        Ec(0x2B617F7D4ED8C7F5L,"1.0000000000002451E-99", "1.0000000000003E-99"),//Ec(0x2B617F7D4ED8C7F5L,"1.000000000000245E-99", "1.0000000000003E-99"),
        Ec(0x2B617F7D4ED8C7F4L,"1.0000000000002449E-99","1.0000000000002E-99"),
        Ec(0x2B617F7D4ED8C609L,"1.0000000000001452E-99","1.0000000000002E-99"),
        Ec(0x2B617F7D4ED8C608L,"1.000000000000145E-99", "1.0000000000001E-99"),
        Ec(0x2B617F7D4ED8C41CL,"1.0000000000000451E-99", "1.0000000000001E-99"),//Ec(0x2B617F7D4ED8C41CL,"1.000000000000045E-99", "1.0000000000001E-99"),
        Ec(0x2B617F7D4ED8C41BL,"1.0000000000000449E-99","1E-99"),
        Ec(0x2B617F7D4ED8C323L,"9.9999999999999454E-100","1E-99"),//Ec(0x2B617F7D4ED8C323L,"9.999999999999945E-100","1E-99"),
        Ec(0x2B617F7D4ED8C322L,"9.9999999999999434E-100","9.9999999999999E-100"),//Ec(0x2B617F7D4ED8C322L,"9.999999999999943E-100","9.9999999999999E-100"),
        Ec(0x2B617F7D4ED8C2F2L,"9.9999999999998459E-100","9.9999999999999E-100"),//Ec(0x2B617F7D4ED8C2F2L,"9.999999999999846E-100","9.9999999999999E-100"),
        Ec(0x2B617F7D4ED8C2F1L,"9.9999999999998439E-100","9.9999999999998E-100"),//Ec(0x2B617F7D4ED8C2F1L,"9.999999999999844E-100","9.9999999999998E-100"),
        Ec(0x2B617F7D4ED8C2C1L,"9.9999999999997464E-100","9.9999999999998E-100"),//Ec(0x2B617F7D4ED8C2C1L,"9.999999999999746E-100","9.9999999999998E-100"),
        Ec(0x2B617F7D4ED8C2C0L,"9.9999999999997444E-100","9.9999999999997E-100"),//Ec(0x2B617F7D4ED8C2C0L,"9.999999999999744E-100","9.9999999999997E-100"),



        // small numbers
        Ec(0x3EE9E409302678BAL, "1.2345678901234568E-05", "1.23456789012346E-05"),//Ec(0x3EE9E409302678BAL, "1.2345678901234568E-5", "1.23456789012346E-05"),
        Ec(0x3F202E85BE180B74L, "0.00012345678901234567", "0.000123456789012346"),//Ec(0x3F202E85BE180B74L, "1.2345678901234567E-4", "0.000123456789012346"),
        Ec(0x3F543A272D9E0E51L, "0.0012345678901234567", "0.00123456789012346"),
        Ec(0x3F8948B0F90591E6L, "0.012345678901234568", "0.0123456789012346"),

        Ec(0x3EE9E409301B5A02L, "1.23456789E-05", "0.0000123456789"),//Ec(0x3EE9E409301B5A02L, "1.23456789E-5", "0.0000123456789"),

        Ec(0x3E6E7D05BDABDE50L, "5.6789012345E-08", "0.000000056789012345"),
        Ec(0x3E6E7D05BDAD407EL, "5.67890123456E-08", "5.67890123456E-08"),//Ec(0x3E6E7D05BDAD407EL, "5.67890123456E-8", "5.67890123456E-08"),
        Ec(0x3E6E7D06029F18BEL, "5.678902E-08", "0.00000005678902"),//Ec(0x3E6E7D06029F18BEL, "5.678902E-8", "0.00000005678902"),

        Ec(0x2BCB5733CB32AE6EL, "9.9999999999991232E-98",  "9.99999999999912E-98"),//Ec(0x2BCB5733CB32AE6EL, "9.999999999999123E-98",  "9.99999999999912E-98"),
        Ec(0x2B617F7D4ED8C59EL, "1.0000000000001235E-99", "1.0000000000001E-99"),
        Ec(0x0036319916D67853L, "1.2345678901234578E-307", "1.2345678901235E-307"),

        Ec(0x359DEE7A4AD4B81FL, "2E-50", "2E-50"),

        // large numbers 
        Ec(0x41678C29DCD6E9E0L, "12345678.901234567", "12345678.9012346"),//Ec(0x41678C29DCD6E9E0L, "1.2345678901234567E+7", "12345678.9012346"),
        Ec(0x42A674E79C5FE523L, "12345678901234.568", "12345678901234.6"),//Ec(0x42A674E79C5FE523L, "1.2345678901234568E+13", "12345678901234.6"),
        Ec(0x42DC12218377DE6BL, "123456789012345.67", "123456789012346"),//Ec(0x42DC12218377DE6BL, "1.2345678901234567E+14", "123456789012346"),
        Ec(0x43118B54F22AEB03L, "1234567890123456.8", "1234567890123460"),//Ec(0x43118B54F22AEB03L, "1.2345678901234568E+15", "1234567890123460"),
        Ec(0x43E56A95319D63E1L, "1.2345678901234567E+19", "12345678901234600000"),
        Ec(0x441AC53A7E04BCDAL, "1.2345678901234568E+20", "1.23456789012346E+20"),
        Ec(0xC3E56A95319D63E1L, "-1.2345678901234567E+19", "-12345678901234600000"),
        Ec(0xC41AC53A7E04BCDAL, "-1.2345678901234568E+20", "-1.23456789012346E+20"),

        Ec(0x54820FE0BA17F46DL, "1.2345678901234577E+99",  "1.2345678901235E+99"),
        Ec(0x54B693D8E89DF188L, "1.2345678901234576E+100", "1.2345678901235E+100"),

        Ec(0x4A611B0EC57E649AL, "2E+50", "2E+50"),

        // range extremities
        Ec(0x7FEFFFFFFFFFFFFFL, "1.7976931348623157E+308", "1.7976931348623E+308"),
        Ec(0x0010000000000000L, "2.2250738585072014E-308", "2.2250738585072E-308"),
        Ec(0x000FFFFFFFFFFFFFL, "2.2250738585072009E-308", "0"),//Ec(0x000FFFFFFFFFFFFFL, "2.225073858507201E-308", "0"),
        Ec(0x0000000000000001L, "4.94065645841247E-324", "0"),//Ec(0x0000000000000001L, "4.9E-324", "0"),

		// infinity
		Ec(0x7FF0000000000000L, "Infinity", "1.7976931348623E+308"),
		Ec(0xFFF0000000000000L, "-Infinity", "1.7976931348623E+308"),

		// shortening due to rounding
		Ec(0x441AC7A08EAD02F2L, "1.234999999999999E+20", "1.235E+20"),
		Ec(0x40FE26BFFFFFFFF9L, "123499.9999999999", "123500"),
		Ec(0x3E4A857BFB2F2809L, "1.234999999999999E-08", "0.00000001235"),//Ec(0x3E4A857BFB2F2809L, "1.234999999999999E-8", "0.00000001235"),
		Ec(0x3BCD291DEF868C89L, "1.234999999999999E-20", "1.235E-20"),

		// carry up due to rounding
		// For clarity these tests choose values that don't round in java,
		// but will round in excel. In some cases there is almost no difference
		// between excel and java (e.g. 9.9..9E-8)
		Ec(0x444B1AE4D6E2EF4FL, "9.9999999999999987E+20", "1E+21"),//Ec(0x444B1AE4D6E2EF4FL, "9.999999999999999E20", "1E+21"),
		Ec(0x412E847FFFFFFFFFL, "999999.99999999988", "1000000"),//Ec(0x412E847FFFFFFFFFL, "999999.9999999999", "1000000"),
		Ec(0x3E45798EE2308C39L, "9.9999999999999986E-09", "0.00000001"),//Ec(0x3E45798EE2308C39L, "9.999999999999999E-9", "0.00000001"),
		Ec(0x3C32725DD1D243ABL, "9.9999999999999988E-19", "0.000000000000000001"),//Ec(0x3C32725DD1D243ABL, "9.999999999999999E-19", "0.000000000000000001"),
		Ec(0x3BFD83C94FB6D2ABL, "9.9999999999999985E-20", "1E-19"),//Ec(0x3BFD83C94FB6D2ABL, "9.999999999999999E-20", "1E-19"),

		Ec(0xC44B1AE4D6E2EF4FL, "-9.9999999999999987E+20", "-1E+21"),//Ec(0xC44B1AE4D6E2EF4FL, "-9.999999999999999E20", "-1E+21"),
		Ec(0xC12E847FFFFFFFFFL, "-999999.99999999988", "-1000000"),//Ec(0xC12E847FFFFFFFFFL, "-999999.9999999999", "-1000000"),
		Ec(0xBE45798EE2308C39L, "-9.9999999999999986E-09", "-0.00000001"),//Ec(0xBE45798EE2308C39L, "-9.999999999999999E-9", "-0.00000001"),
		Ec(0xBC32725DD1D243ABL, "-9.9999999999999988E-19", "-0.000000000000000001"),//Ec(0xBC32725DD1D243ABL, "-9.999999999999999E-19", "-0.000000000000000001"),
		Ec(0xBBFD83C94FB6D2ABL, "-9.9999999999999985E-20", "-1E-19"),//Ec(0xBBFD83C94FB6D2ABL, "-9.999999999999999E-20", "-1E-19"),


		// NaNs
		// Currently these test cases are not critical, since other limitations prevent any variety in
		// or control of the bit patterns used to encode NaNs in evaluations.
		Ec(0xFFFF0420003C0000L, "NaN", "3.484840871308E+308"),
		Ec(0x7FF8000000000000L, "NaN", "2.6965397022935E+308"),
		Ec(0x7FFF0420003C0000L, "NaN", "3.484840871308E+308"),
		Ec(0xFFF8000000000000L, "NaN", "2.6965397022935E+308"),
		Ec(0xFFFF0AAAAAAAAAAAL, "NaN", "3.4877119413344E+308"),
		Ec(0x7FF80AAAAAAAAAAAL, "NaN", "2.7012211948322E+308"),
		Ec(0xFFFFFFFFFFFFFFFFL, "NaN", "3.5953862697246E+308"),
		Ec(0x7FFFFFFFFFFFFFFFL, "NaN", "3.5953862697246E+308"),
		Ec(0xFFF7FFFFFFFFFFFFL, "NaN", "2.6965397022935E+308"),
	};

        private static ExampleConversion Ec(long rawDoubleBits, String javaRendering, String excelRendering)
        {
            return new ExampleConversion(rawDoubleBits, javaRendering, excelRendering);
        }
        private static ExampleConversion Ec(ulong rawDoubleBits, String javaRendering, String excelRendering)
        {
            return new ExampleConversion((long)rawDoubleBits, javaRendering, excelRendering);
        }
        public static ExampleConversion[] GetExampleConversions()
        {
            return (ExampleConversion[])examples.Clone();
        }
    }

}