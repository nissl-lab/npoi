using System;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.SS.Util
{
/**
 * Helper methods for when working with Usermodel Workbooks
 */
public class WorkbookUtil {

    /**
     * Creates a valid sheet name, which is conform to the rules.
     * In any case, the result safely can be used for 
     * {@link org.apache.poi.ss.usermodel.Workbook#setSheetName(int, String)}.
     * <br/>
     * Rules:
     * <ul>
     * <li>never null</li>
     * <li>minimum length is 1</li>
     * <li>maximum length is 31</li>
     * <li>doesn't contain special chars: 0x0000, 0x0003, / \ ? * ] [ </li>
     * <li>Sheet names must not begin or end with ' (apostrophe)</li>
     * </ul>
     * Invalid characters are replaced by one space character ' '.
     * 
     * @param nameProposal can be any string, will be truncated if necessary,
     *        allowed to be null
     * @return a valid string, "empty" if to short, "null" if null         
     */
    public static String CreateSafeSheetName(String nameProposal)
    {
        return CreateSafeSheetName(nameProposal, ' ');
	}
    /**
     * Creates a valid sheet name, which is conform to the rules.
     * In any case, the result safely can be used for
     * {@link org.apache.poi.ss.usermodel.Workbook#setSheetName(int, String)}.
     * <br />
     * Rules:
     * <ul>
     * <li>never null</li>
     * <li>minimum length is 1</li>
     * <li>maximum length is 31</li>
     * <li>doesn't contain special chars: : 0x0000, 0x0003, / \ ? * ] [ </li>
     * <li>Sheet names must not begin or end with ' (apostrophe)</li>
     * </ul>
     *
     * @param nameProposal can be any string, will be truncated if necessary,
     *        allowed to be null
     * @param replaceChar the char to replace invalid characters.
     * @return a valid string, "empty" if to short, "null" if null
     */
    public static String CreateSafeSheetName(String nameProposal, char replaceChar) {
        if (nameProposal == null) {
            return "null";
        }
        if (nameProposal.Length < 1) {
            return "empty";
        }
        int length = Math.Min(31, nameProposal.Length);
        String shortenname = nameProposal.Substring(0, length);
        StringBuilder result = new StringBuilder(shortenname);
        for (int i=0; i<length; i++) {
            char ch = result[(i)];
            switch (ch) {
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
                    if (i==0 || i==length-1) {
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
    /**
     * Validates sheet name.
     *
     * <p>
     * The character count <c>MUST</c> be greater than or equal to 1 and less than or equal to 31.
     * The string MUST NOT contain the any of the following characters:
     * <ul>
     * <li> 0x0000 </li>
     * <li> 0x0003 </li>
     * <li> colon (:) </li>
     * <li> backslash (\) </li>
     * <li> asterisk (*) </li>
     * <li> question mark (?) </li>
     * <li> forward slash (/) </li>
     * <li> opening square bracket ([) </li>
     * <li> closing square bracket (]) </li>
     * </ul>
     * The string MUST NOT begin or end with the single quote (') character.
     * </p>
     *
     * @param sheetName the name to validate
     */
    public static void ValidateSheetName(String sheetName) {
        if (sheetName == null) {
            throw new ArgumentException("sheetName must not be null");
        }
        int len = sheetName.Length;
        if (len < 1 || len > 31) {
            throw new ArgumentException("sheetName '" + sheetName
                    + "' is invalid - character count MUST be greater than or equal to 1 and less than or equal to 31");
        }

        for (int i=0; i<len; i++) {
            char ch = sheetName[i];
            switch (ch) {
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
        if (sheetName[0] == '\'' || sheetName[len-1] == '\'') {
            throw new ArgumentException("Invalid sheet name '" + sheetName
                    + "'. Sheet names must not begin or end with (').");
        }
    }


    public static void ValidateSheetState(SheetState state)
    {
        switch(state){
            case SheetState.Visible: break;
            case SheetState.Hidden: break;
            case SheetState.VeryHidden: break;
            default: throw new ArgumentException("Ivalid sheet state : " + state + "\n" +
                            "Sheet state must beone of the Workbook.SHEET_STATE_* constants");
        }
    }    
}

}
