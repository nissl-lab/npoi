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
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using NPOI.OpenXml4Net.Exceptions;
using System.Collections;

namespace NPOI.OpenXml4Net.OPC.Internal
{
    /// <summary>
    /// <para>
    /// Represents a immutable MIME ContentType value (RFC 2616 &#167;3.7)<br/>
    /// media-type = type "/" subtype *( ";" parameter ) type = token<br/>
    /// subtype = token<br/>
    /// </para>
    /// <para>
    /// Rule M1.13 : Package implementers shall only create and only recognize parts
    /// with a content type; format designers shall specify a content type for each
    /// part included in the format. Content types for package parts shall fit the
    /// definition and syntax for media types as specified in RFC 2616, \&#167;3.7.
    /// Rule M1.14: Content types shall not use linear white space either between the
    /// type and subtype or between an attribute and its value. Content types also
    /// shall not have leading or trailing white spaces. Package implementers shall
    /// create only such content types and shall require such content types when
    /// retrieving a part from a package; format designers shall specify only such
    /// content types for inclusion in the format.
    /// </para>
    /// </summary>
    /// 
    /// see <a href="http://www.ietf.org/rfc/rfc2045.txt">http://www.ietf.org/rfc/rfc2045.txt</a>
    /// see <a href="http://www.ietf.org/rfc/rfc2616.txt">http://www.ietf.org/rfc/rfc2616.txt</a>
    /// <remarks>
    /// @author Julien Chable
    /// @version 0.1
    /// </remarks>

    public class ContentType:IComparable
    {

        /// <summary>
        /// Type in Type/Subtype.
        /// </summary>
        private String type;

        /// <summary>
        /// Subtype
        /// </summary>
        private String subType;

        /// <summary>
        /// Parameters
        /// </summary>
        private Dictionary<String, String> parameters;

        /// <summary>
        /// Media type compiled pattern for parameters.
        /// </summary>
        private static Regex patternTypeSubType;

        /// <summary>
        /// Media type compiled pattern, with parameters.
        /// </summary>
        private static Regex patternTypeSubTypeParams;
        /// <summary>
        /// Pattern to match on just the parameters part, to work
        /// around the Java Regexp group capture behaviour
        /// </summary>
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

        /// <summary>
        /// Constructor. Check the input with the RFC 2616 grammar.
        /// </summary>
        /// <param name="contentType">The content type to store.
        /// </param>
        /// <exception cref="InvalidFormatException">InvalidFormatException
        /// If the specified content type is not valid with RFC 2616.
        /// </exception>
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
            else
            {
                // missing media type and subtype
                this.type = "";
                this.subType = "";
                this.parameters = new Dictionary<string, string>();
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
                foreach (KeyValuePair<string, string> kv in parameters)
                {
                    retVal.Append(";");
                    retVal.Append(kv.Key);
                    retVal.Append("=");
                    retVal.Append(kv.Value);
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

        /// <summary>
        /// Get the subtype.
        /// </summary>
        /// <returns>The subtype of this content type.</returns>
        public String SubType
        {
            get
            {
                return this.subType;
            }
        }

        /// <summary>
        /// Get the type.
        /// </summary>
        /// <returns>The type of this content type.</returns>
        public String Type
        {
            get
            {
                return this.type;
            }
        }

        /// <summary>
        /// Does this content type have any parameters associated with it?
        /// </summary>
        public bool HasParameters()
        {
            return (parameters != null) && !(parameters.Count == 0);
        }

        /// <summary>
        /// Return the parameter keys
        /// </summary>
        public String[] GetParameterKeys()
        {
            if (parameters == null)
                return new String[0];
            List<string> keys = new List<string>();
            keys.AddRange(parameters.Keys);
            return keys.ToArray();
        }

        /// <summary>
        /// Gets the value associated to the specified key.
        /// </summary>
        /// <param name="key">The key of the key/value pair.
        /// </param>
        /// <returns>The value associated to the specified key.</returns>
        public String GetParameter(String key)
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