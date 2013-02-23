using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using NPOI.OpenXml4Net.Exceptions;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /**
     * Represents a immutable MIME ContentType value (RFC 2616 &#167;3.7)
     * <p>
     * media-type = type "/" subtype *( ";" parameter ) type = token<br>
     * subtype = token<br>
     * </p><p>
     * Rule M1.13 : Package implementers shall only create and only recognize parts
     * with a content type; format designers shall specify a content type for each
     * part included in the format. Content types for package parts shall fit the
     * definition and syntax for media types as specified in RFC 2616, \&#167;3.7.
     * </p><p>
     * Rule M1.14: Content types shall not use linear white space either between the
     * type and subtype or between an attribute and its value. Content types also
     * shall not have leading or trailing white spaces. Package implementers shall
     * create only such content types and shall require such content types when
     * retrieving a part from a package; format designers shall specify only such
     * content types for inclusion in the format.
     * </p>
     * @author Julien Chable
     * @version 0.1
     *
     * @see <a href="http://www.ietf.org/rfc/rfc2045.txt">http://www.ietf.org/rfc/rfc2045.txt</a>
     * @see <a href="http://www.ietf.org/rfc/rfc2616.txt">http://www.ietf.org/rfc/rfc2616.txt</a>
     */
    public class ContentType:IComparable
    {

        /**
         * Type in Type/Subtype.
         */
        private String type;

        /**
         * Subtype
         */
        private String subType;

        /**
         * Parameters
         */
        private SortedList<String, String> parameters;

        /**
         * Media type compiled pattern for parameters.
         */
        private static Regex patternMediaType;

        static ContentType()
        {
            /*
             * token = 1*<any CHAR except CTLs or separators>
             *
             * separators = "(" | ")" | "<" | ">" | "@" | "," | ";" | ":" | "\" |
             * <"> | "/" | "[" | "]" | "?" | "=" | "{" | "}" | SP | HT
             *
             * CTL = <any US-ASCII control character (octets 0 - 31) and DEL (127)>
             *
             * CHAR = <any US-ASCII character (octets 0 - 127)>
             */
            String token = "[^\\(\\)<>@,;:\\\\/\"\\[\\]\\?={}\\s]";

            /*
             * parameter = attribute "=" value
             *
             * attribute = token
             *
             * value = token | quoted-string
             */
            // Keep for future use with parameter:
            // String parameter = "(" + token + "+)=(\"?" + token + "+\"?)";
            /*
             * Pattern for media type.
             *
             * Don't allow comment, rule M1.15: The package implementer shall
             * require a content type that does not include comments and the format
             * designer shall specify such a content type.
             *
             * comment = "(" *( ctext | quoted-pair | comment ) ")"
             *
             * ctext = <any TEXT excluding "(" and ")">
             *
             * TEXT = <any OCTET except CTLs, but including LWS>
             *
             * LWS = [CRLF] 1*( SP | HT )
             *
             * CR = <US-ASCII CR, carriage return (13)>
             *
             * LF = <US-ASCII LF, linefeed (10)>
             *
             * SP = <US-ASCII SP, space (32)>
             *
             * HT = <US-ASCII HT, horizontal-tab (9)>
             *
             * quoted-pair = "\" CHAR
             */

            // Keep for future use with parameter:
            // patternMediaType = Pattern.compile("^(" + token + "+)/(" + token
            // + "+)(;" + parameter + ")*$");
            patternMediaType = new Regex("^(" + token + "+)/(" + token
                    + "+)$");
        }

        /**
         * Constructor. Check the input with the RFC 2616 grammar.
         *
         * @param contentType
         *            The content type to store.
         * @throws InvalidFormatException
         *             If the specified content type is not valid with RFC 2616.
         */
        public ContentType(String contentType)
        {
            Match mMediaType = patternMediaType.Match(contentType);
            if (!mMediaType.Success)
                throw new InvalidFormatException(
                        "The specified content type '"
                                + contentType
                                + "' is not compliant with RFC 2616: malformed content type.");

            // Type/subtype
            if (mMediaType.Groups.Count >= 2)
            {
                this.type = mMediaType.Groups[1].Value;
                this.subType = mMediaType.Groups[2].Value;
                // Parameters
                this.parameters = new SortedList<String, String>();
                for (int i = 4; i <= mMediaType.Groups.Count
                        && (mMediaType.Groups[i] != null); i += 2)
                {
                    this.parameters[mMediaType.Groups[i].Value] = mMediaType
                            .Groups[i + 1].Value;
                }
            }
        }


        public override String ToString()
        {
            StringBuilder retVal = new StringBuilder();
            retVal.Append(this.Type);
            retVal.Append("/");
            retVal.Append(this.SubType);
            // Keep for future implementation if needed
            // for (String key : parameters.keySet()) {
            // retVal.Append(";");
            // retVal.Append(key);
            // retVal.Append("=");
            // retVal.Append(parameters.get(key));
            // }
            return retVal.ToString();
        }


        public override bool Equals(Object obj)
        {
            return (!(obj is ContentType))
                    || (this.ToString().Equals(obj.ToString(), StringComparison.InvariantCultureIgnoreCase));
        }

        


        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        /* Getters */

        /**
         * Get the subtype.
         *
         * @return The subtype of this content type.
         */
        public String SubType
        {
            get
            {
                return this.subType;
            }
        }

        /**
         * Get the type.
         *
         * @return The type of this content type.
         */
        public String Type
        {
            get
            {
                return this.type;
            }
        }

        /**
         * Gets the value associated to the specified key.
         *
         * @param key
         *            The key of the key/value pair.
         * @return The value associated to the specified key.
         */
        public String GetParameters(String key)
        {
            return parameters[key];
        }

        #region IComparable Members

        public int CompareTo(object obj)
        {
            if (obj == null)
                return -1;

            if (this.Equals(obj))
                return 0;

            return 1;
        }

        #endregion
    }
}