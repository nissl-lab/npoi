using System;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.SS.Util
{
    /// <summary>
    /// Helper methods for when working with Usermodel Workbooks
    /// </summary>
    public class WorkbookUtil
    {

        /// <summary>
        /// Creates a valid sheet name, which is conform to the rules.
        /// In any case, the result safely can be used for
        /// <see cref="org.apache.poi.ss.usermodel.Workbook.setSheetName(int, String)" />.
        /// <br/>
        /// Rules:
        /// <list type="bullet">
        /// <item><description>never null</description></item>
        /// <item><description>minimum length is 1</description></item>
        /// <item><description>maximum length is 31</description></item>
        /// <item><description>doesn't contain special chars: 0x0000, 0x0003, / \ ? * ] [ </description></item>
        /// <item><description>Sheet names must not begin or end with ' (apostrophe)</description></item>
        /// </list>
        /// Invalid characters are replaced by one space character ' '.
        /// </summary>
        /// <param name="nameProposal">can be any string, will be truncated if necessary,
        /// allowed to be null
        /// </param>
        /// <return>a valid string, "empty" if to short, "null" if null</return>
        public static String CreateSafeSheetName(String nameProposal)
        {
            return CreateSafeSheetName(nameProposal, ' ');
        }
        /// <summary>
        /// Creates a valid sheet name, which is conform to the rules.
        /// In any case, the result safely can be used for
        /// <see cref="org.apache.poi.ss.usermodel.Workbook.setSheetName(int, String)" />.
        /// <br />
        /// Rules:
        /// <list type="bullet">
        /// <item><description>never null</description></item>
        /// <item><description>minimum length is 1</description></item>
        /// <item><description>maximum length is 31</description></item>
        /// <item><description>doesn't contain special chars: : 0x0000, 0x0003, / \ ? * ] [ </description></item>
        /// <item><description>Sheet names must not begin or end with ' (apostrophe)</description></item>
        /// </list>
        /// </summary>
        /// <param name="nameProposal">can be any string, will be truncated if necessary,
        /// allowed to be null
        /// </param>
        /// <param name="replaceChar">the char to replace invalid characters.</param>
        /// <return>a valid string, "empty" if to short, "null" if null</return>
        public static String CreateSafeSheetName(String nameProposal, char replaceChar)
        {
            if(nameProposal == null)
            {
                return "null";
            }
            if(nameProposal.Length < 1)
            {
                return "empty";
            }
            int length = Math.Min(31, nameProposal.Length);
            String shortenname = nameProposal.Substring(0, length);
            StringBuilder result = new StringBuilder(shortenname);
            for(int i = 0; i<length; i++)
            {
                char ch = result[(i)];
                switch(ch)
                {
                    case '\u0000':
                    case '\u0003':
                    case ':':
                    case '/':
                    case '\\':
                    case '?':
                    case '*':
                    case ']':
                    case '[':
                        result[i] = replaceChar;
                        break;
                    case '\'':
                        if(i==0 || i==length-1)
                        {
                            result[i] = replaceChar;
                        }
                        break;
                    default:
                        // all other chars OK
                        break;
                }
            }
            return result.ToString();
        }
        /// <summary>
        /// <para>
        /// Validates sheet name.
        /// </para>
        /// <para>
        /// 
        /// The character count <c>MUST</c> be greater than or equal to 1 and less than or equal to 31.
        /// The string MUST NOT contain the any of the following characters:
        /// <list type="bullet">
        /// <item><description> 0x0000 </description></item>
        /// <item><description> 0x0003 </description></item>
        /// <item><description> colon (:) </description></item>
        /// <item><description> backslash (\) </description></item>
        /// <item><description> asterisk (*) </description></item>
        /// <item><description> question mark (?) </description></item>
        /// <item><description> forward slash (/) </description></item>
        /// <item><description> opening square bracket ([) </description></item>
        /// <item><description> closing square bracket (]) </description></item>
        /// </list>
        /// The string MUST NOT begin or end with the single quote (') character.
        /// </para>
        /// </summary>
        /// <param name="sheetName">the name to validate</param>
        public static void ValidateSheetName(String sheetName)
        {
            if(sheetName == null)
            {
                throw new ArgumentException("sheetName must not be null");
            }
            int len = sheetName.Length;
            if(len < 1 || len > 31)
            {
                throw new ArgumentException("sheetName '" + sheetName
                        + "' is invalid - character count MUST be greater than or equal to 1 and less than or equal to 31");
            }

            for(int i = 0; i<len; i++)
            {
                char ch = sheetName[i];
                switch(ch)
                {
                    case '/':
                    case '\\':
                    case '?':
                    case '*':
                    case ']':
                    case '[':
                    case ':':
                        break;
                    default:
                        // all other chars OK
                        continue;
                }
                throw new ArgumentException("Invalid char (" + ch
                        + ") found at index (" + i + ") in sheet name '" + sheetName + "'");
            }
            if(sheetName[0] == '\'' || sheetName[len-1] == '\'')
            {
                throw new ArgumentException("Invalid sheet name '" + sheetName
                        + "'. Sheet names must not begin or end with (').");
            }
        }

        public static void ValidateSheetState(SheetVisibility state)
        {
            switch(state)
            {
                case SheetVisibility.Visible:
                    break;
                case SheetVisibility.Hidden:
                    break;
                case SheetVisibility.VeryHidden:
                    break;
                default:
                    throw new ArgumentException("Invalid sheet state : " + state + "\n" +
                                "Sheet state must beone of the Workbook.SHEET_STATE_* constants");
            }
        }

        public static int GetNextActiveSheetDueToSheetHiding(IWorkbook wb, int sheetIx)
        {
            if(sheetIx == wb.ActiveSheetIndex)
            {
                // activate next sheet
                // if last sheet in workbook, the previous visible sheet should be activated
                int count = wb.NumberOfSheets;
                for(int i = sheetIx+1; i < count; i++)
                {
                    // get the next visible sheet in this workbook
                    if(SheetVisibility.Visible == wb.GetSheetVisibility(i))
                    {
                        return i;
                    }
                }

                // if there are no sheets to the right or all sheets to the right are hidden, activate a sheet to the left
                for(int i = sheetIx-1; i < count; i--)
                {
                    if(SheetVisibility.Visible == wb.GetSheetVisibility(i))
                    {
                        return i;
                    }
                }

                // there are no other visible sheets in this workbook
                return -1;
                //throw new IllegalStateException("Cannot hide sheet " + sheetIx + ". Workbook must contain at least 1 other visible sheet.");
            }
            else
            {
                return sheetIx;
            }
        }
    }

}
