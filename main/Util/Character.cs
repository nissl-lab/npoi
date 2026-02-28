namespace NPOI.Util
{
    public class Character
    {
        public static int GetNumericValue(char src)
        {
            if (src >= '0' && src <= '9')
            {
                return (int)src;
            }
            if (src >= 'A' && src <= 'Z')
            {
                return ((int)src) - 55;
            }
            if (src >= 'a' && src <= 'z')
            {
                return ((int)src) - 87;
            }
            return -1;
        }

        //TODO: this should work but maybe not.
        public static bool isWhitespace(char src)
        {
            return char.IsWhiteSpace(src);
        }
    }
}
