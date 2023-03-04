using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace NPOI.SS.Formula.Functions
{
    public class DateValue : Fixed1ArgFunction
    {
        private class Format
        {
            public Regex pattern;
            public bool hasYear;
            public int yearIndex;
            public int monthIndex;
            public int dayIndex;

            public Format(string patternString, String groupOrder)
            {
                this.pattern = new Regex(patternString, RegexOptions.Compiled);
                this.hasYear = groupOrder.Contains("y");
                if (hasYear)
                {
                    yearIndex = groupOrder.IndexOf("y");
                }
                monthIndex = groupOrder.IndexOf("m");
                dayIndex = groupOrder.IndexOf("d");
            }
            private static List<Format> formats = new List<Format>();
            static Format()
            {
                formats.Add(new Format("^(\\d{4})-(\\w+)-(\\d{1,2})$", "ymd"));
                formats.Add(new Format("^(\\d{1,2})-(\\w+)-(\\d{4})$", "dmy"));
                formats.Add(new Format("^(\\w+)-(\\d{1,2})$", "md"));
                formats.Add(new Format("^(\\w+)/(\\d{1,2})/(\\d{4})$", "mdy"));
                formats.Add(new Format("^(\\d{4})/(\\w+)/(\\d{1,2})$", "ymd"));
                formats.Add(new Format("^(\\w+)/(\\d{1,2})$", "md"));
            }
            public static List<Format> Values()
            {
                return formats;
            }
        }

        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval dateTextArg)
        {
            try
            {
                String dateText = OperandResolver.CoerceValueToString(
                OperandResolver.GetSingleValue(dateTextArg, srcRowIndex, srcColumnIndex));
                if (string.IsNullOrEmpty(dateText))
                {
                    return BlankEval.instance;
                }
                foreach (Format format in Format.Values())
                {
                    var m = format.pattern.Match(dateText);
                    if (m.Success)
                    {
                        var matchGroups = m.Groups;
                        List<String> groups = new List<string>();
                        for (int i = 1; i <= matchGroups.Count; ++i)
                        {
                            groups.Add(matchGroups[i].Value);
                        }
                        int year = format.hasYear
                            ? int.Parse(groups[format.yearIndex])
                            : DateTime.Now.Year;
                        int month = parseMonth(groups[format.monthIndex]);
                        int day = int.Parse(groups[format.dayIndex]);
                        return new NumberEval(DateUtil.GetExcelDate(new DateTime(year, month, day)));
                    }
                }
            }
            catch (FormatException)
            {
                return ErrorEval.VALUE_INVALID;
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            return ErrorEval.VALUE_INVALID;
        }
        private int parseMonth(String monthPart)
        {
            try
            {
                return int.Parse(monthPart);
            }
            catch (FormatException)
            {
            }

            string[] months = DateTimeFormatInfo.InvariantInfo.MonthNames;
            for (int month = 0; month < months.Length; ++month)
            {
                if (months[month].ToLower().StartsWith(monthPart.ToLower()))
                {
                    return month + 1;
                }
            }
            return -1;
        }
    }
}
