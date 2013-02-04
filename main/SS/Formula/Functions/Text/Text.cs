using System;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;

namespace NPOI.SS.Formula.Functions
{
    /**
	 * An implementation of the TEXT function
	 * TEXT returns a number value formatted with the given number formatting string. 
	 * This function is not a complete implementation of the Excel function, but
	 *  handles most of the common cases. All work is passed down to 
	 *  {@link DataFormatter} to be done, as this works much the same as the
	 *  display focused work that that does. 
	 */
    public class Text : Fixed2ArgFunction
    {
        public static DataFormatter Formatter = new DataFormatter();
        public override ValueEval Evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1)
        {
            double s0;
            String s1;
            try
            {
                s0 = TextFunction.EvaluateDoubleArg(arg0, srcRowIndex, srcColumnIndex);
                s1 = TextFunction.EvaluateStringArg(arg1, srcRowIndex, srcColumnIndex);
            }
            catch (EvaluationException e)
            {
                return e.GetErrorEval();
            }
            try
            {
                // Ask DataFormatter to handle the String for us
                String formattedStr = Formatter.FormatRawCellContents(s0, -1, s1);
                return new StringEval(formattedStr);
            }
            catch (Exception)
            {
                return ErrorEval.VALUE_INVALID;
            }
            //if (Regex.Match(s1, "[y|m|M|d|s|h]+").Success)
            //{
            //    //may be datetime string
            //    ValueEval result = TryParseDateTime(s0, s1);
            //    if (result != ErrorEval.VALUE_INVALID)
            //        return result;
            //}
            ////The regular expression needs ^ and $. 
            //if (Regex.Match(s1, @"^[\d,\#,\.,\$,\,]+$").Success)
            //{
            //    //TODO: simulate DecimalFormat class in java.
            //    FormatBase formatter = new DecimalFormat(s1);
            //    return new StringEval(formatter.Format(s0, CultureInfo.CurrentCulture));
            //}
            //else if (s1.IndexOf("/", StringComparison.Ordinal) == s1.LastIndexOf("/", StringComparison.Ordinal) && s1.IndexOf("/", StringComparison.Ordinal) >= 0 && !s1.Contains("-"))
            //{
            //    double wholePart = Math.Floor(s0);
            //    double decPart = s0 - wholePart;
            //    if (wholePart * decPart == 0)
            //    {
            //        return new StringEval("0");
            //    }
            //    String[] parts = s1.Split(' ');
            //    String[] fractParts;
            //    if (parts.Length == 2)
            //    {
            //        fractParts = parts[1].Split('/');
            //    }
            //    else
            //    {
            //        fractParts = s1.Split('/');
            //    }

            //    if (fractParts.Length == 2)
            //    {
            //        double minVal = 1.0;
            //        double currDenom = Math.Pow(10, fractParts[1].Length) - 1d;
            //        double currNeum = 0;
            //        for (int i = (int)(Math.Pow(10, fractParts[1].Length) - 1d); i > 0; i--)
            //        {
            //            for (int i2 = (int)(Math.Pow(10, fractParts[1].Length) - 1d); i2 > 0; i2--)
            //            {
            //                if (minVal >= Math.Abs((double)i2 / (double)i - decPart))
            //                {
            //                    currDenom = i;
            //                    currNeum = i2;
            //                    minVal = Math.Abs((double)i2 / (double)i - decPart);
            //                }
            //            }
            //        }
            //        FormatBase neumFormatter = new DecimalFormat(fractParts[0]);
            //        FormatBase denomFormatter = new DecimalFormat(fractParts[1]);
            //        if (parts.Length == 2)
            //        {
            //            FormatBase wholeFormatter = new DecimalFormat(parts[0]);
            //            String result = wholeFormatter.Format(wholePart, CultureInfo.CurrentCulture) + " " + neumFormatter.Format(currNeum, CultureInfo.CurrentCulture) + "/" + denomFormatter.Format(currDenom, CultureInfo.CurrentCulture);
            //            return new StringEval(result);
            //        }
            //        else
            //        {
            //            String result = neumFormatter.Format(currNeum + (currDenom * wholePart), CultureInfo.CurrentCulture) + "/" + denomFormatter.Format(currDenom, CultureInfo.CurrentCulture);
            //            return new StringEval(result);
            //        }
            //    }
            //    else
            //    {
            //        return ErrorEval.VALUE_INVALID;
            //    }
            //}
            //else
            //{
            //    return TryParseDateTime(s0, s1);
            //}
        }
        //private ValueEval TryParseDateTime(double s0, string s1)
        //{
        //    try
        //    {
        //        FormatBase dateFormatter = new SimpleDateFormat(s1);
        //        //first month of java Gregorian Calendar month field is 0
        //        DateTime dt = new DateTime(1899, 12, 30, 0, 0, 0);
        //        dt = dt.AddDays((int)Math.Floor(s0));
        //        double dayFraction = s0 - Math.Floor(s0);
        //        dt = dt.AddMilliseconds((int)Math.Round(dayFraction * 24 * 60 * 60 * 1000));
        //        return new StringEval(dateFormatter.Format(dt, CultureInfo.CurrentCulture));
        //    }
        //    catch (Exception)
        //    {
        //        return ErrorEval.VALUE_INVALID;
        //    }
        //}
    }
}