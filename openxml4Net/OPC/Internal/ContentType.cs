using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using NPOI.OpenXml4Net.Exceptions;
using System.Collections;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /**
     * Represents a immutable MIME ContentType value (RFC 2616 &#167;3.7)
     * media-type = type "/" subtype *( ";" parameter ) type = token<br>
     * subtype = token<br>

     * Rule M1.13 : Package implementers shall only create and only recognize parts
     * with a content type; format designers shall specify a content type for each
     * part included in the format. Content types for package parts shall fit the
     * definition and syntax for media types as specified in RFC 2616, \&#167;3.7.
     * Rule M1.14: Content types shall not use linear white space either between the
     * type and subtype or between an attribute and its value. Content types also
     * shall not have leading or trailing white spaces. Package implementers shall
     * create only such content types and shall require such content types when
     * retrieving a part from a package; format designers shall specify only such
     * content types for inclusion in the format.
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
        Hashtable p;
        private Dictionary<String, String> parameters;

        /**
         * Media type compiled pattern for parameters.
         */
        private static Regex patternTypeSubType;

        /**
         * Media type compiled pattern, with parameters.
         */
        private static Regex patternTypeSubTypeParams;
        /**
         * Pattern to match on just the parameters part, to work
         * around the Java Regexp group capture behaviour
         */
        private static Regex patternParams;
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
            //String token = "[\\x21-\\x7E&&[^\\(\\)<>@,;:\\\\/\"\\[\\]\\?={}\\x20\\x09]]";
            string token = @"[\x21\x23-\x27\x2A\x2B\x2D\x2E0-9A-Z\x5E\x5F\x60a-z\x7E]";

            /*
             * parameter = attribute "=" value
             *
             * attribute = token
             *
             * value = token | quoted-string
             */
            String parameter = "(" + token + "+)=(\"?" + token + "+\"?)";
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

            patternTypeSubType = new Regex("^(" + token + "+)/(" + token + "+)$");
            patternTypeSubTypeParams = new Regex("^(" + token + "+)/(" + token + "+)(;" + parameter + ")*$");
            patternParams = new Regex(";" + parameter);
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
            Match mMediaType = patternTypeSubType.Match(contentType);
            if (!mMediaType.Success)
                // How about with parameters?
                mMediaType = patternTypeSubTypeParams.Match(contentType);
            if (!mMediaType.Success)
            {
                throw new InvalidFormatException(
                        "The specified content type '"
                                + contentType
                                + "' is not compliant with RFC 2616: malformed content type.");
            }
            // Type/subtype
            if (mMediaType.Groups.Count >= 2)
            {
                this.type = mMediaType.Groups[1].Value;
                this.subType = mMediaType.Groups[2].Value;
                // Parameters
                this.parameters = new Dictionary<String, String>();
                // Java RegExps are unhelpful, and won't do multiple group captures
                // See http://docs.oracle.com/javase/6/docs/api/java/util/regex/Pattern.html#cg
                if (mMediaType.Groups.Count >= 5)
                {
                    Match mParams = patternParams.Match(contentType.Substring(mMediaType.Groups[2].Index + mMediaType.Groups[2].Length));
                    while (mParams.Success)
                    {
                        this.parameters.Add(mParams.Groups[1].Value, mParams.Groups[2].Value);
                        mParams = mParams.NextMatch();
                    }
                }
            }
        }
        public override String ToString()
        {
            return ToString(true);
        }
        public String ToString(bool withParameters)
        {
            StringBuilder retVal = new StringBuilder();
            retVal.Append(this.Type);
            retVal.Append("/");
            retVal.Append(this.SubType);
            if (withParameters)
            {
                foreach (String key in parameters.Keys)
                {
                    retVal.Append(";");
                    retVal.Append(key);
                    retVal.Append("=");
                    retVal.Append(parameters[key]);
                }
            }
            return retVal.ToString();
        }
        public String ToStringWithParameters()
        {
            StringBuilder retVal = new StringBuilder();
            retVal.Append(ToString());

            foreach (String key in parameters.Keys)
            {
                retVal.Append(";");
                retVal.Append(key);
                retVal.Append("=");
                retVal.Append(parameters[(key)]);
            }
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
         * Does this content type have any parameters associated with it?
         */
        public bool HasParameters()
        {
            return (parameters != null) && !(parameters.Count == 0);
        }

        /**
         * Return the parameter keys
         */
        public String[] GetParameterKeys()
        {
            if (parameters == null)
                return new String[0];
            List<string> keys = new List<string>();
            keys.AddRange(parameters.Keys);
            return keys.ToArray();
        }

        /**
         * Gets the value associated to the specified key.
         *
         * @param key
         *            The key of the key/value pair.
         * @return The value associated to the specified key.
         */
        public String GetParameter(String key)
        {
            return parameters[key];
        }

        /**
     * @deprecated Use {@link #getParameter(String)} instead
     */
        public String GetParameters(String key)
        {
            return GetParameter(key);
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