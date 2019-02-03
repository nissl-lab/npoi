using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using NPOI.OpenXml4Net.Exceptions;
using NPOI.Util;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * An immutable Open Packaging Convention compliant part name.
     *
     * @author Julien Chable
     *
     * @see <a href="http://www.ietf.org/rfc/rfc3986.txt">http://www.ietf.org/rfc/rfc3986.txt</a>
     */
    public class PackagePartName : IComparable<PackagePartName>
    {

        /**
         * Part name stored as an URI.
         */
        private Uri partNameURI;

        /*
         * URI Characters definition (RFC 3986)
         */

        /**
         * Reserved characters for sub delimitations.
         */
        private static String[] RFC3986_PCHAR_SUB_DELIMS = { "!", "$", "&", "'",
            "(", ")", "*", "+", ",", ";", "=" };

        /**
         * Unreserved character (+ ALPHA & DIGIT).
         */
        private static String[] RFC3986_PCHAR_UNRESERVED_SUP = { "-", ".", "_", "~" };

        /**
         * Authorized reserved characters for pChar.
         */
        private static String[] RFC3986_PCHAR_AUTHORIZED_SUP = { ":", "@" };

        /**
         * Flag to know if this part name is from a relationship part name.
         */
        private bool isRelationship;

        /**
         * Constructor. Makes a ValidPartName object from a java.net.URI
         *
         * @param uri
         *            The URI to validate and to transform into ValidPartName.
         * @param checkConformance
         *            Flag to specify if the contructor have to validate the OPC
         *            conformance. Must be always <code>true</code> except for
         *            special URI like '/' which is needed for internal use by
         *            OpenXml4Net but is not valid.
         * @throws InvalidFormatException
         *             Throw if the specified part name is not conform to Open
         *             Packaging Convention specifications.
         * @see java.net.URI
         */
        public PackagePartName(Uri uri, bool checkConformance)
        {
            if (checkConformance)
            {
                ThrowExceptionIfInvalidPartUri(uri);
            }
            else
            {
                if (!PackagingUriHelper.PACKAGE_ROOT_URI.Equals(uri))
                {
                    throw new OpenXml4NetException(
                            "OCP conformance must be check for ALL part name except special cases : ['/']");
                }
            }
            this.partNameURI = uri;
            this.isRelationship = IsRelationshipPartURI(this.partNameURI);
        }

        /**
         * Constructor. Makes a ValidPartName object from a String part name.
         *
         * @param partName
         *            Part name to valid and to create.
         * @param checkConformance
         *            Flag to specify if the contructor have to validate the OPC
         *            conformance. Must be always <code>true</code> except for
         *            special URI like '/' which is needed for internal use by
         *            OpenXml4Net but is not valid.
         * @throws InvalidFormatException
         *             Throw if the specified part name is not conform to Open
         *             Packaging Convention specifications.
         */
        internal PackagePartName(String partName, bool checkConformance)
        {
            Uri partURI;
            try
            {
                partURI = PackagingUriHelper.ParseUri(partName, UriKind.RelativeOrAbsolute);
            }
            catch (UriFormatException)
            {
                throw new ArgumentException(
                        "partName argmument is not a valid OPC part name !");
            }

            if (checkConformance)
            {
                ThrowExceptionIfInvalidPartUri(partURI);
            }
            else
            {
                if (!PackagingUriHelper.PACKAGE_ROOT_URI.Equals(partURI))
                {
                    throw new OpenXml4NetException(
                            "OCP conformance must be check for ALL part name except special cases : ['/']");
                }
            }
            this.partNameURI = partURI;
            this.isRelationship = IsRelationshipPartURI(this.partNameURI);
        }

        /**
         * Check if the specified part name is a relationship part name.
         *
         * @param partUri
         *            The URI to check.
         * @return <code>true</code> if this part name respect the relationship
         *         part naming convention else <code>false</code>.
         */
        private bool IsRelationshipPartURI(Uri partUri)
        {
            if (partUri == null)
                throw new ArgumentException("partUri");

            return Regex.IsMatch(partUri.OriginalString,
                    "^.*/" + PackagingUriHelper.RELATIONSHIP_PART_SEGMENT_NAME + "/.*\\"
                            + PackagingUriHelper.RELATIONSHIP_PART_EXTENSION_NAME
                            + "$");
        }

        /**
         * Know if this part name is a relationship part name.
         *
         * @return <code>true</code> if this part name respect the relationship
         *         part naming convention else <code>false</code>.
         */
        public bool IsRelationshipPartURI()
        {
            return this.isRelationship;
        }

        /**
         * Throws an exception (of any kind) if the specified part name does not
         * follow the Open Packaging Convention specifications naming rules.
         *
         * @param partUri
         *            The part name to check.
         * @throws Exception
         *             Throws if the part name is invalid.
         */
        private static void ThrowExceptionIfInvalidPartUri(Uri partUri)
        {
            if (partUri == null)
                throw new ArgumentException("partUri");
            // Check if the part name URI is empty [M1.1]
            ThrowExceptionIfEmptyURI(partUri);

            // Check if the part name URI is absolute
            ThrowExceptionIfAbsoluteUri(partUri);

            // Check if the part name URI starts with a forward slash [M1.4]
            ThrowExceptionIfPartNameNotStartsWithForwardSlashChar(partUri);

            // Check if the part name URI ends with a forward slash [M1.5]
            ThrowExceptionIfPartNameEndsWithForwardSlashChar(partUri);

            // Check if the part name does not have empty segments. [M1.3]
            // Check if a segment ends with a dot ('.') character. [M1.9]
            ThrowExceptionIfPartNameHaveInvalidSegments(partUri);
        }

        /**
         * Throws an exception if the specified URI is empty. [M1.1]
         *
         * @param partURI
         *            Part URI to check.
         * @throws InvalidFormatException
         *             If the specified URI is empty.
         */
        private static void ThrowExceptionIfEmptyURI(Uri partURI)
        {
            if (partURI == null)
                throw new ArgumentException("partURI");

            String uriPath = partURI.OriginalString;
            if (uriPath.Length == 0
                    || ((uriPath.Length == 1) && (uriPath[0] == PackagingUriHelper.FORWARD_SLASH_CHAR)))
                throw new InvalidFormatException(
                        "A part name shall not be empty [M1.1]: "
                                + partURI.OriginalString);
        }

        /**
         * Throws an exception if the part name has empty segments. [M1.3]
         *
         * Throws an exception if a segment any characters other than pchar
         * characters. [M1.6]
         *
         * Throws an exception if a segment contain percent-encoded forward slash
         * ('/'), or backward slash ('\') characters. [M1.7]
         *
         * Throws an exception if a segment contain percent-encoded unreserved
         * characters. [M1.8]
         *
         * Throws an exception if the specified part name's segments end with a dot
         * ('.') character. [M1.9]
         *
         * Throws an exception if a segment doesn't include at least one non-dot
         * character. [M1.10]
         *
         * @param partUri
         *            The part name to check.
         * @throws InvalidFormatException
         *             if the specified URI contain an empty segments or if one the
         *             segments contained in the part name, ends with a dot ('.')
         *             character.
         */
        private static void ThrowExceptionIfPartNameHaveInvalidSegments(Uri partUri)
        {
            if (partUri == null || "".Equals(partUri))
            {
                throw new ArgumentException("partUri");
            }

            // Split the URI into several part and analyze each
            String[] segments = partUri.OriginalString.Split('/');
            if (segments.Length <= 1 || !segments[0].Equals(""))
                throw new InvalidFormatException(
                        "A part name shall not have empty segments [M1.3]: "
                                + partUri.OriginalString);

            for (int i = 1; i < segments.Length; ++i)
            {
                String seg = segments[i];
                if (seg == null || "".Equals(seg))
                {
                    throw new InvalidFormatException(
                            "A part name shall not have empty segments [M1.3]: "
                                    + partUri.OriginalString);
                }

                if (seg.EndsWith("."))
                {
                    throw new InvalidFormatException(
                            "A segment shall not end with a dot ('.') character [M1.9]: "
                                    + partUri.OriginalString);
                }

                if ("".Equals(seg.Replace("\\\\.", "")))
                {
                    // Normally will never been invoked with the previous
                    // implementation rule [M1.9]
                    throw new InvalidFormatException(
                            "A segment shall include at least one non-dot character. [M1.10]: "
                                    + partUri.OriginalString);
                }

                // Check for rule M1.6, M1.7, M1.8
                CheckPCharCompliance(seg);
            }
        }

        /**
         * Throws an exception if a segment any characters other than pchar
         * characters. [M1.6]
         *
         * Throws an exception if a segment contain percent-encoded forward slash
         * ('/'), or backward slash ('\') characters. [M1.7]
         *
         * Throws an exception if a segment contain percent-encoded unreserved
         * characters. [M1.8]
         *
         * @param segment
         *            The segment to check
         */
        private static void CheckPCharCompliance(String segment)
        {
            bool errorFlag;
            for (int i = 0; i < segment.Length; ++i)
            {
                char c = segment[i];
                errorFlag = true;

                /* Check rule M1.6 */

                // Check for digit or letter
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')
                        || (c >= '0' && c <= '9'))
                {
                    errorFlag = false;
                }
                else
                {
                    // Check "-", ".", "_", "~"
                    for (int j = 0; j < RFC3986_PCHAR_UNRESERVED_SUP.Length; ++j)
                    {
                        if (c == RFC3986_PCHAR_UNRESERVED_SUP[j][0])
                        {
                            errorFlag = false;
                            break;
                        }
                    }

                    // Check ":", "@"
                    for (int j = 0; errorFlag
                            && j < RFC3986_PCHAR_AUTHORIZED_SUP.Length; ++j)
                    {
                        if (c == RFC3986_PCHAR_AUTHORIZED_SUP[j][0])
                        {
                            errorFlag = false;
                        }
                    }

                    // Check "!", "$", "&", "'", "(", ")", "*", "+", ",", ";", "="
                    for (int j = 0; errorFlag
                            && j < RFC3986_PCHAR_SUB_DELIMS.Length; ++j)
                    {
                        if (c == RFC3986_PCHAR_SUB_DELIMS[j][0])
                        {
                            errorFlag = false;
                        }
                    }
                }

                if (errorFlag && c == '%')
                {
                    // We certainly found an encoded character, check for length
                    // now ( '%' HEXDIGIT HEXDIGIT)
                    if (((segment.Length - i) < 2))
                    {
                        throw new InvalidFormatException("The segment " + segment
                                + " contain invalid encoded character !");
                    }

                    // If not percent encoded character error occur then reset the
                    // flag -> the character is valid
                    errorFlag = false;

                    // Decode the encoded character
                    char decodedChar = (char)Convert.ToInt32(segment.Substring(
                            i + 1, 2), 16);
                    i += 2;

                    /* Check rule M1.7 */
                    if (decodedChar == '/' || decodedChar == '\\')
                        throw new InvalidFormatException(
                                "A segment shall not contain percent-encoded forward slash ('/'), or backward slash ('\') characters. [M1.7]");

                    /* Check rule M1.8 */

                    // Check for unreserved character like define in RFC3986
                    if ((decodedChar >= 'A' && decodedChar <= 'Z')
                            || (decodedChar >= 'a' && decodedChar <= 'z')
                            || (decodedChar >= '0' && decodedChar <= '9'))
                        errorFlag = true;

                    // Check for unreserved character "-", ".", "_", "~"
                    for (int j = 0; !errorFlag
                            && j < RFC3986_PCHAR_UNRESERVED_SUP.Length; ++j)
                    {
                        if (c == RFC3986_PCHAR_UNRESERVED_SUP[j][0])
                        {
                            errorFlag = true;
                            break;
                        }
                    }
                    if (errorFlag)
                        throw new InvalidFormatException(
                                "A segment shall not contain percent-encoded unreserved characters. [M1.8]");
                }

                if (errorFlag)
                    throw new InvalidFormatException(
                            "A segment shall not hold any characters other than pchar characters. [M1.6]");
            }
        }

        /**
         * Throws an exception if the specified part name doesn't start with a
         * forward slash character '/'. [M1.4]
         *
         * @param partUri
         *            The part name to check.
         * @throws InvalidFormatException
         *             If the specified part name doesn't start with a forward slash
         *             character '/'.
         */
        private static void ThrowExceptionIfPartNameNotStartsWithForwardSlashChar(
                Uri partUri)
        {
            String uriPath = partUri.OriginalString;
            if (uriPath.Length > 0
                    && uriPath[0] != PackagingUriHelper.FORWARD_SLASH_CHAR)
                throw new InvalidFormatException(
                        "A part name shall start with a forward slash ('/') character [M1.4]: "
                                + partUri.OriginalString);
        }

        /**
         * Throws an exception if the specified part name ends with a forwar slash
         * character '/'. [M1.5]
         *
         * @param partUri
         *            The part name to check.
         * @throws InvalidFormatException
         *             If the specified part name ends with a forwar slash character
         *             '/'.
         */
        private static void ThrowExceptionIfPartNameEndsWithForwardSlashChar(
                Uri partUri)
        {
            String uriPath = partUri.OriginalString;
            if (uriPath.Length > 0
                    && uriPath[uriPath.Length - 1] == PackagingUriHelper.FORWARD_SLASH_CHAR)
                throw new InvalidFormatException(
                        "A part name shall not have a forward slash as the last character [M1.5]: "
                                + partUri.OriginalString);
        }

        /**
         * Throws an exception if the specified URI is absolute.
         *
         * @param partUri
         *            The URI to check.
         * @throws InvalidFormatException
         *             Throws if the specified URI is absolute.
         */
        private static void ThrowExceptionIfAbsoluteUri(Uri partUri)
        {
            if (partUri.IsAbsoluteUri)
                throw new InvalidFormatException("Absolute URI forbidden: "
                        + partUri);
        }

        /**
         * Compare two part name following the rule M1.12 :
         *
         * Part name equivalence is determined by comparing part names as
         * case-insensitive ASCII strings. Packages shall not contain equivalent
         * part names and package implementers shall neither create nor recognize
         * packages with equivalent part names. [M1.12]
         */
        public int CompareTo(PackagePartName other)
        {
            // compare with natural sort order
            return Compare(this, other);
        }

        /**
         * Retrieves the extension of the part name if any. If there is no extension
         * returns an empty String. Example : '/document/content.xml' => 'xml'
         *
         * @return The extension of the part name.
         */
        public String Extension
        {
            get
            {
                String fragment = this.partNameURI.OriginalString;
                if (fragment.Length > 0)
                {
                    int i = fragment.LastIndexOf(".");
                    if (i > -1)
                        return fragment.Substring(i + 1);
                }
                return "";
            }
        }

        /**
         * Get this part name.
         *
         * @return The name of this part name.
         */
        public String Name
        {
            get
            {
                return this.partNameURI.OriginalString;
            }
        }

        /**
         * Part name equivalence is determined by comparing part names as
         * case-insensitive ASCII strings. Packages shall not contain equivalent
         * part names and package implementers shall neither create nor recognize
         * packages with equivalent part names. [M1.12]
         */

        public override bool Equals(Object other)
        {
            if (other is PackagePartName)
            {
                // String.equals() is compatible with our compareTo(), but cheaper
                return this.partNameURI.OriginalString.ToLower().Equals
                (
                    ((PackagePartName)other).partNameURI.OriginalString.ToLower()
                );
            }
            else
            {
                return false;
            }
        }


        public override int GetHashCode()
        {
            return this.partNameURI.OriginalString.ToLower().GetHashCode();
        }


        public override String ToString()
        {
            return this.Name;
        }

        /* Getters and setters */

        /**
         * Part name property getter.
         *
         * @return This part name URI.
         */
        public Uri URI
        {
            get
            {
                return this.partNameURI;
            }
        }

        /**
         * A natural sort order for package part names, consistent with the
         * requirements of {@code java.util.Comparator}, but simply implemented
         * as a static method.
         * <p>
         * For example, this sorts "file10.png" after "file2.png" (comparing the
         * numerical portion), but sorts "File10.png" before "file2.png"
         * (lexigraphical sort)
         *
         * <p>
         * When comparing part names, the rule M1.12 is followed:
         *
         * Part name equivalence is determined by comparing part names as
         * case-insensitive ASCII strings. Packages shall not contain equivalent
         * part names and package implementers shall neither create nor recognize
         * packages with equivalent part names. [M1.12]
         */
        public static int Compare(PackagePartName obj1, PackagePartName obj2)
        {
            // NOTE could also throw a NullPointerException() if desired
            if (obj1 == null)
            {
                // (null) == (null), (null) < (non-null)
                return (obj2 == null ? 0 : -1);
            }
            else if (obj2 == null)
            {
                // (non-null) > (null)
                return 1;
            }

            return Compare
            (
                obj1.URI.OriginalString.ToLower(),
                obj2.URI.OriginalString.ToLower()
            );
        }


        /**
         * A natural sort order for strings, consistent with the
         * requirements of {@code java.util.Comparator}, but simply implemented
         * as a static method.
         * <p>
         * For example, this sorts "file10.png" after "file2.png" (comparing the
         * numerical portion), but sorts "File10.png" before "file2.png"
         * (lexigraphical sort)
         */
        public static int Compare(String str1, String str2)
        {
            if (str1 == null)
            {
                // (null) == (null), (null) < (non-null)
                return (str2 == null ? 0 : -1);
            }
            else if (str2 == null)
            {
                // (non-null) > (null)
                return 1;
            }

            int len1 = str1.Length;
            int len2 = str2.Length;
            for (int idx1 = 0, idx2 = 0; idx1 < len1 && idx2 < len2; /*nil*/)
            {
                char c1 = str1[(idx1++)];
                char c2 = str2[(idx2++)];

                if (char.IsDigit(c1) && char.IsDigit(c2))
                {
                    int beg1 = idx1 - 1;  // undo previous increment
                    while (idx1 < len1 && char.IsDigit(str1[(idx1)]))
                    {
                        ++idx1;
                    }

                    int beg2 = idx2 - 1;  // undo previous increment
                    while (idx2 < len2 && char.IsDigit(str2[(idx2)]))
                    {
                        ++idx2;
                    }

                    // note: BigInteger for extra safety
                    //int cmp = new BigInteger(str1.Substring(beg1, idx1 - beg1)).CompareTo
                    //(
                    //    new BigInteger(str2.Substring(beg2, idx2 - beg2))
                    //);
                    int cmp = decimal.Parse(str1.Substring(beg1, idx1 - beg1)).CompareTo(
                        decimal.Parse(str2.Substring(beg2, idx2 - beg2))
                        );
                    if (cmp != 0) return cmp;
                }
                else if (c1 != c2)
                {
                    return (c1 - c2);
                }
            }

            return (len1 - len2);
        }
    }
}