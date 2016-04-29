using System;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.OpenXml4Net.Exceptions;
using System.IO;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Helper for part and pack Uri.
     *
     * @author Julien Chable, CDubet, Kim Ung
     * @version 0.1
     */
    public class PackagingUriHelper
    {

        /**
         * Package root Uri.
         */
        private static Uri packageRootUri;

        /**
         * Extension name of a relationship part.
         */
        public static String RELATIONSHIP_PART_EXTENSION_NAME;

        /**
         * Segment name of a relationship part.
         */
        public static String RELATIONSHIP_PART_SEGMENT_NAME;

        /**
         * Segment name of the package properties folder.
         */
        public static String PACKAGE_PROPERTIES_SEGMENT_NAME;

        /**
         * Core package properties art name.
         */
        public static String PACKAGE_CORE_PROPERTIES_NAME;

        /**
         * Forward slash Uri separator.
         */
        public static char FORWARD_SLASH_CHAR;

        /**
         * Forward slash Uri separator.
         */
        public static String FORWARD_SLASH_STRING;

        /**
         * Package relationships part Uri
         */
        public static Uri PACKAGE_RELATIONSHIPS_ROOT_URI;

        /**
         * Package relationships part name.
         */
        public static PackagePartName PACKAGE_RELATIONSHIPS_ROOT_PART_NAME;

        /**
         * Core properties part Uri.
         */
        public static Uri CORE_PROPERTIES_URI;

        /**
         * Core properties partname.
         */
        public static PackagePartName CORE_PROPERTIES_PART_NAME;

        /**
         * Root package Uri.
         */
        public static Uri PACKAGE_ROOT_URI;

        /**
         * Root package part name.
         */
        public static PackagePartName PACKAGE_ROOT_PART_NAME;

        /* Static initialization */
        static PackagingUriHelper()
        {
            RELATIONSHIP_PART_SEGMENT_NAME = "_rels";
            RELATIONSHIP_PART_EXTENSION_NAME = ".rels";
            FORWARD_SLASH_CHAR = '/';
            FORWARD_SLASH_STRING = "/";
            PACKAGE_PROPERTIES_SEGMENT_NAME = "docProps";
            PACKAGE_CORE_PROPERTIES_NAME = "core.xml";

            // Make Uri
            Uri uriPACKAGE_ROOT_URI = null;
            Uri uriPACKAGE_RELATIONSHIPS_ROOT_URI = null;
            Uri uriPACKAGE_PROPERTIES_URI = null;

            uriPACKAGE_ROOT_URI = ParseUri("/",UriKind.Relative);
            uriPACKAGE_RELATIONSHIPS_ROOT_URI = ParseUri(FORWARD_SLASH_CHAR
                    + RELATIONSHIP_PART_SEGMENT_NAME + FORWARD_SLASH_CHAR
                    + RELATIONSHIP_PART_EXTENSION_NAME, UriKind.Relative);
            packageRootUri = ParseUri("/", UriKind.Relative);
            uriPACKAGE_PROPERTIES_URI = ParseUri(FORWARD_SLASH_CHAR
                    + PACKAGE_PROPERTIES_SEGMENT_NAME + FORWARD_SLASH_CHAR
                    + PACKAGE_CORE_PROPERTIES_NAME, UriKind.Relative);

            PACKAGE_ROOT_URI = uriPACKAGE_ROOT_URI;
            PACKAGE_RELATIONSHIPS_ROOT_URI = uriPACKAGE_RELATIONSHIPS_ROOT_URI;
            CORE_PROPERTIES_URI = uriPACKAGE_PROPERTIES_URI;

            // Make part name from previous Uri
            PackagePartName tmpPACKAGE_ROOT_PART_NAME = null;
            PackagePartName tmpPACKAGE_RELATIONSHIPS_ROOT_PART_NAME = null;
            PackagePartName tmpCORE_PROPERTIES_URI = null;
            try
            {
                tmpPACKAGE_RELATIONSHIPS_ROOT_PART_NAME = CreatePartName(PACKAGE_RELATIONSHIPS_ROOT_URI);
                tmpCORE_PROPERTIES_URI = CreatePartName(CORE_PROPERTIES_URI);
                tmpPACKAGE_ROOT_PART_NAME = new PackagePartName(PACKAGE_ROOT_URI,
                        false);
            }
            catch (InvalidFormatException)
            {
                // Should never happen in production as all data are fixed
            }
            PACKAGE_RELATIONSHIPS_ROOT_PART_NAME = tmpPACKAGE_RELATIONSHIPS_ROOT_PART_NAME;
            CORE_PROPERTIES_PART_NAME = tmpCORE_PROPERTIES_URI;
            PACKAGE_ROOT_PART_NAME = tmpPACKAGE_ROOT_PART_NAME;
        }
        private static Regex missingAuthPattern = new Regex("\\w+://$");
        /**
         * Gets the Uri for the package root.
         *
         * @return Uri of the package root.
         */
        public static Uri PackageRootUri
        {
            get
            {
                return packageRootUri;
            }
        }

        private static readonly bool IsMono = Type.GetType("Mono.Runtime") != null;
        public static Uri ParseUri(string s, UriKind kind)
        {
            if (IsMono)
            {
                if (kind == UriKind.Absolute)
                    throw new UriFormatException();
                if (kind == UriKind.RelativeOrAbsolute && s.StartsWith("/"))
                    kind = UriKind.Relative;
            }
            return new Uri(s, kind);
        }


        /**
         * Know if the specified Uri is a relationship part name.
         *
         * @param partUri
         *            Uri to check.
         * @return <i>true</i> if the Uri <i>false</i>.
         */
        public static bool IsRelationshipPartURI(Uri partUri)
        {
            if (partUri == null)
                throw new ArgumentException("partUri");

            return Regex.IsMatch(partUri.OriginalString,
                    ".*" + RELATIONSHIP_PART_SEGMENT_NAME + ".*"
                            + RELATIONSHIP_PART_EXTENSION_NAME + "$");
        }

        /**
         * Get file name from the specified Uri.
         */
        public static String GetFilename(Uri uri)
        {
            if (uri != null)
            {
                String path = uri.OriginalString;
                int len = path.Length;
                int num2 = len;
                while (--num2 >= 0)
                {
                    char ch1 = path[num2];
                    if (ch1 == PackagingUriHelper.FORWARD_SLASH_CHAR)
                        return path.Substring(num2 + 1);
                }
            }
            return "";
        }

        /**
         * Get the file name without the trailing extension.
         */
        public static String GetFilenameWithoutExtension(Uri uri)
        {
            String filename = GetFilename(uri);
            int dotIndex = filename.LastIndexOf(".");
            if (dotIndex == -1)
                return filename;
            return filename.Substring(0, dotIndex);
        }

        /**
         * Get the directory path from the specified Uri.
         */
        public static Uri GetPath(Uri uri)
        {
            if (uri != null)
            {
                String path = uri.OriginalString;
                int len = path.Length;
                int num2 = len;
                while (--num2 >= 0)
                {
                    char ch1 = path[num2];
                    if (ch1 == PackagingUriHelper.FORWARD_SLASH_CHAR)
                    {
                        try
                        {
                            return ParseUri(path.Substring(0, num2), UriKind.Absolute);
                        }
                        catch (UriFormatException)
                        {
                            return null;
                        }
                    }
                }
            }
            return null;
        }

        /**
         * Combine two URIs.
         *
         * @param prefix the prefix Uri
         * @param suffix the suffix Uri
         *
         * @return the Combined Uri
         */
        public static Uri Combine(Uri prefix, Uri suffix)
        {
            Uri retUri = null;
            try
            {
                retUri = ParseUri(Combine(prefix.OriginalString, suffix.OriginalString), UriKind.Absolute);
            }
            catch (UriFormatException)
            {
                throw new ArgumentException(
                        "Prefix and suffix can't be Combine !");
            }
            return retUri;
        }

        /**
         * Combine a string Uri with a prefix and a suffix.
         */
        public static String Combine(String prefix, String suffix)
        {
            if (!prefix.EndsWith("" + FORWARD_SLASH_CHAR)
                    && !suffix.StartsWith("" + FORWARD_SLASH_CHAR))
                return prefix + FORWARD_SLASH_CHAR + suffix;
            else if ((!prefix.EndsWith("" + FORWARD_SLASH_CHAR)
                    && suffix.StartsWith("" + FORWARD_SLASH_CHAR) || (prefix
                    .EndsWith("" + FORWARD_SLASH_CHAR) && !suffix.StartsWith(""
                    + FORWARD_SLASH_CHAR))))
                return prefix + suffix;
            else
                return "";
        }

        /**
         * Fully relativize the source part Uri against the target part Uri.
         *
         * @param sourceURI
         *            The source part Uri.
         * @param targetURI
         *            The target part Uri.
         * @return A fully relativize part name Uri ('word/media/image1.gif',
         *         '/word/document.xml' => 'media/image1.gif') else
         *         <code>null</code>.
         */
        public static Uri RelativizeUri(Uri sourceURI, Uri targetURI, bool msCompatible)
        {
            StringBuilder retVal = new StringBuilder();
            String[] segmentsSource = sourceURI.ToString().Split(new char[] { '/' });
            String[] segmentsTarget = targetURI.ToString().Split(new char[] { '/' });

            // If the source Uri is empty
            if (segmentsSource.Length == 0)
            {
                throw new ArgumentException(
                        "Can't relativize an empty source Uri !");
            }

            // If target Uri is empty
            if (segmentsTarget.Length == 0)
            {
                throw new ArgumentException(
                        "Can't relativize an empty target Uri !");
            }

            // If the source is the root, then the relativized
            //  form must actually be an absolute Uri
            if (sourceURI.ToString().Equals("/"))
            {
                String path = targetURI.ToString();
                if (msCompatible && path.Length > 0 && path[0] == '/')
                {
                    try
                    {
                        targetURI = ParseUri(path.Substring(1), UriKind.RelativeOrAbsolute);
                    }
                    catch (Exception)
                    {
                        //_logger.log(POILogger.WARN, e);
                        return null;
                    }
                }
                return targetURI;
            }

            // Relativize the source Uri against the target Uri.
            // First up, figure out how many steps along we can go
            // and still have them be the same
            int segmentsTheSame = 0;
            for (int i = 0; i < segmentsSource.Length && i < segmentsTarget.Length; i++)
            {
                if (segmentsSource[i].Equals(segmentsTarget[i]))
                {
                    // Match so far, good
                    segmentsTheSame++;
                }
                else
                {
                    break;
                }
            }

            // If we didn't have a good match or at least except a first empty element
            if ((segmentsTheSame == 0 || segmentsTheSame == 1) &&
                    segmentsSource[0].Equals("") && segmentsTarget[0].Equals(""))
            {
                for (int i = 0; i < segmentsSource.Length - 2; i++)
                {
                    retVal.Append("../");
                }
                for (int i = 0; i < segmentsTarget.Length; i++)
                {
                    if (segmentsTarget[i].Equals(""))
                        continue;
                    retVal.Append(segmentsTarget[i]);
                    if (i != segmentsTarget.Length - 1)
                        retVal.Append("/");
                }

                try
                {
                    return ParseUri(retVal.ToString(),UriKind.RelativeOrAbsolute);
                }
                catch (Exception)
                {
                    //System.err.println(e);
                    return null;
                }
            }

            // Special case for where the two are the same
            if (segmentsTheSame == segmentsSource.Length
                    && segmentsTheSame == segmentsTarget.Length)
            {
                if (sourceURI.Equals(targetURI))
                {
                    // if source and target are the same they should be resolved to the last segment,
                    // Example: if a slide references itself, e.g. the source URI is
                    // "/ppt/slides/slide1.xml" and the targetURI is "slide1.xml" then
                    // this it should be relativized as "slide1.xml", i.e. the last segment.
                    retVal.Append(segmentsSource[segmentsSource.Length - 1]);
                }
                else
                {
                    retVal.Append("");
                }
            }
            else
            {
                // Matched for so long, but no more

                // Do we need to go up a directory or two from
                // the source to get here?
                // (If it's all the way up, then don't bother!)
                if (segmentsTheSame == 1)
                {
                    retVal.Append("/");
                }
                else
                {
                    for (int j = segmentsTheSame; j < segmentsSource.Length - 1; j++)
                    {
                        retVal.Append("../");
                    }
                }

                // Now go from here on down
                for (int j = segmentsTheSame; j < segmentsTarget.Length; j++)
                {
                    if (retVal.Length > 0
                            && retVal[retVal.Length - 1] != '/')
                    {
                        retVal.Append("/");
                    }
                    retVal.Append(segmentsTarget[j]);
                }
            }

            try
            {
                return ParseUri(retVal.ToString(), UriKind.RelativeOrAbsolute);
            }
            catch (Exception)
            {
                //System.err.println(e);
                return null;
            }
        }


        /**
         * Fully relativize the source part URI against the target part URI.
         *
         * @param sourceURI
         *            The source part URI.
         * @param targetURI
         *            The target part URI.
         * @return A fully relativize part name URI ('word/media/image1.gif',
         *         '/word/document.xml' => 'media/image1.gif') else
         *         <code>null</code>.
         */
        public static Uri RelativizeUri(Uri sourceURI, Uri targetURI)
        {
            return RelativizeUri(sourceURI, targetURI, false);
        }

        /**
         * Resolve a source uri against a target.
         *
         * @param sourcePartUri
         *            The source Uri.
         * @param targetUri
         *            The target Uri.
         * @return The resolved Uri.
         */
        public static Uri ResolvePartUri(Uri sourcePartUri, Uri targetUri)
        {
            if (sourcePartUri == null || sourcePartUri.IsAbsoluteUri)
            {
                throw new ArgumentException("sourcePartUri invalid - "
                        + sourcePartUri);
            }

            if (targetUri == null || targetUri.IsAbsoluteUri)
            {
                throw new ArgumentException("targetUri invalid - "
                        + targetUri);
            }
            string path;
            if (sourcePartUri.OriginalString == "/")
                path = "/";
            else
                path = Path.GetDirectoryName(sourcePartUri.OriginalString).Replace("\\", "/");

            string targetPath = targetUri.OriginalString;
            if (targetPath.StartsWith("../"))
            {
                string[] segments = path.Split(new char[] { '/' });

                int segmentEnd = segments.Length - 1;
                while (targetPath.StartsWith("../"))
                {
                    targetPath = targetPath.Substring(3);
                    segmentEnd -= 1;
                }
                path = "/";

                for (int i = 0; i <= segmentEnd;i++ )
                {
                    if(segments[i]!=string.Empty)
                    path += segments[i]+"/";
                }
                path += targetPath;
            }
            else
            {
                path = Path.Combine(path, targetUri.OriginalString).Replace("\\", "/");
            }
            return ParseUri(path, UriKind.RelativeOrAbsolute);
        }

        /**
         * Get Uri from a string path.
         */
        public static Uri GetURIFromPath(String path)
        {
            Uri retUri = null;
            try
            {
                retUri = ParseUri(path,UriKind.RelativeOrAbsolute);
            }
            catch (UriFormatException)
            {
                throw new ArgumentException("path");
            }
            return retUri;
        }

        /**
         * Get the source part Uri from a specified relationships part.
         *
         * @param relationshipPartUri
         *            The relationship part use to retrieve the source part.
         * @return The source part Uri from the specified relationships part.
         */
        public static Uri GetSourcePartUriFromRelationshipPartUri(
                Uri relationshipPartUri)
        {
            if (relationshipPartUri == null)
                throw new ArgumentException(
                        "Must not be null");

            if (!IsRelationshipPartURI(relationshipPartUri))
                throw new ArgumentException(
                        "Must be a relationship part");

            if (Uri.Compare(relationshipPartUri, PACKAGE_RELATIONSHIPS_ROOT_URI, UriComponents.AbsoluteUri, UriFormat.SafeUnescaped, StringComparison.InvariantCultureIgnoreCase) == 0)
                return PACKAGE_ROOT_URI;

            String filename = relationshipPartUri.OriginalString;
            String filenameWithoutExtension = GetFilenameWithoutExtension(relationshipPartUri);
            filename = filename
                    .Substring(0, ((filename.Length - filenameWithoutExtension
                            .Length) - RELATIONSHIP_PART_EXTENSION_NAME.Length));
            filename = filename.Substring(0, filename.Length
                    - RELATIONSHIP_PART_SEGMENT_NAME.Length - 1);
            filename = Combine(filename, filenameWithoutExtension);
            return GetURIFromPath(filename);
        }

        /**
         * Create an OPC compliant part name by throwing an exception if the Uri is
         * not valid.
         *
         * @param partUri
         *            The part name Uri to validate.
         * @return A valid part name object, else <code>null</code>.
         * @throws InvalidFormatException
         *             Throws if the specified Uri is not OPC compliant.
         */
        public static PackagePartName CreatePartName(Uri partUri)
        {
            if (partUri == null)
                throw new ArgumentException("partName");

            return new PackagePartName(partUri, true);
        }

        /**
         * Create an OPC compliant part name.
         *
         * @param partName
         *            The part name to validate.
         * @return The correspondant part name if valid, else <code>null</code>.
         * @throws InvalidFormatException
         *             Throws if the specified part name is not OPC compliant.
         * @see #CreatePartName(Uri)
         */
        public static PackagePartName CreatePartName(String partName)
        {
            Uri partNameURI;
            try
            {
                partName = partName.Replace("\\","/");  //tolerate backslash - poi test49609
                partNameURI = ParseUri(partName,UriKind.Relative);
            }
            catch (UriFormatException e)
            {
                throw new InvalidFormatException(e.Message);
            }
            return CreatePartName(partNameURI);
        }

        /**
         * Create an OPC compliant part name by resolving it using a base part.
         *
         * @param partName
         *            The part name to validate.
         * @param relativePart
         *            The relative base part.
         * @return The correspondant part name if valid, else <code>null</code>.
         * @throws InvalidFormatException
         *             Throws if the specified part name is not OPC compliant.
         * @see #CreatePartName(Uri)
         */
        public static PackagePartName CreatePartName(String partName,
                PackagePart relativePart)
        {
            Uri newPartNameURI;
            try
            {
                newPartNameURI = ResolvePartUri(
                        relativePart.PartName.URI, ParseUri(partName,UriKind.RelativeOrAbsolute));
            }
            catch (UriFormatException e)
            {
                throw new InvalidFormatException(e.Message);
            }
            return CreatePartName(newPartNameURI);
        }

        /**
         * Create an OPC compliant part name by resolving it using a base part.
         *
         * @param partName
         *            The part name Uri to validate.
         * @param relativePart
         *            The relative base part.
         * @return The correspondant part name if valid, else <code>null</code>.
         * @throws InvalidFormatException
         *             Throws if the specified part name is not OPC compliant.
         * @see #CreatePartName(Uri)
         */
        public static PackagePartName CreatePartName(Uri partName,
                PackagePart relativePart)
        {
            Uri newPartNameURI = ResolvePartUri(
                    relativePart.PartName.URI, partName);
            return CreatePartName(newPartNameURI);
        }

        /**
         * Validate a part Uri by returning a bool.
         * ([M1.1],[M1.3],[M1.4],[M1.5],[M1.6])
         *
         * (OPC Specifications 8.1.1 Part names) :
         *
         * Part Name Syntax
         *
         * The part name grammar is defined as follows:
         *
         * <i>part_name = 1*( "/" segment )
         *
         * segment = 1*( pchar )</i>
         *
         *
         * (pchar is defined in RFC 3986)
         *
         * @param partUri
         *            The Uri to validate.
         * @return <b>true</b> if the Uri is valid to the OPC Specifications, else
         *         <b>false</b>
         *
         * @see #CreatePartName(Uri)
         */
        public static bool IsValidPartName(Uri partUri)
        {
            if (partUri == null)
                throw new ArgumentException("partUri");

            try
            {
                CreatePartName(partUri);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /**
         * Decode a Uri by converting all percent encoded character into a String
         * character.
         *
         * @param uri
         *            The Uri to decode.
         * @return The specified Uri in a String with converted percent encoded
         *         characters.
         */
        public static String DecodeURI(Uri uri)
        {
            StringBuilder retVal = new StringBuilder();
            String uriStr = uri.OriginalString;
            char c;
            for (int i = 0; i < uriStr.Length; ++i)
            {
                c = uriStr[i];
                if (c == '%')
                {
                    // We certainly found an encoded character, check for length
                    // now ( '%' HEXDIGIT HEXDIGIT)
                    if (((uriStr.Length - i) < 2))
                    {
                        throw new ArgumentException("The uri " + uriStr
                                + " contain invalid encoded character !");
                    }

                    // Decode the encoded character
                    char decodedChar = (char)Convert.ToInt32(uriStr.Substring(
                            i + 1, i + 3), 16);
                    retVal.Append(decodedChar);
                    i += 2;
                    continue;
                }
                retVal.Append(c);
            }
            return retVal.ToString();
        }
           /**
     * Convert a string to {@link java.net.URI}
     *
     * If  part name is not a valid URI, it is resolved as follows:
     * <p>
     * 1. Percent-encode each open bracket ([) and close bracket (]).</li>
     * 2. Percent-encode each percent (%) character that is not followed by a hexadecimal notation of an octet value.</li>
     * 3. Un-percent-encode each percent-encoded unreserved character.
     * 4. Un-percent-encode each forward slash (/) and back slash (\).
     * 5. Convert all back slashes to forward slashes.
     * 6. If present in a segment containing non-dot (?.?) characters, remove trailing dot (?.?) characters from each segment.
     * 7. Replace each occurrence of multiple consecutive forward slashes (/) with a single forward slash.
     * 8. If a single trailing forward slash (/) is present, remove that trailing forward slash.
     * 9. Remove complete segments that consist of three or more dots.
     * 10. Resolve the relative reference against the base URI of the part holding the Unicode string, as it is defined
     * in ?5.2 of RFC 3986. The path component of the resulting absolute URI is the part name.
     *</p>
     *
     * @param   value   the string to be parsed into a URI
     * @return  the resolved part name that should be OK to construct a URI
     *
     * TODO YK: for now this method does only (5). Finish the rest.
     */
        public static Uri ToUri(String value)
        {
            //5. Convert all back slashes to forward slashes
            if (value.IndexOf("\\") != -1)
            {
                value = value.Replace('\\', '/');
            }

            // URI fragemnts (those starting with '#') are not encoded
            // and may contain white spaces and raw unicode characters
            int fragmentIdx = value.IndexOf('#');
            if (fragmentIdx != -1)
            {
                String path = value.Substring(0, fragmentIdx);
                String fragment = value.Substring(fragmentIdx + 1);

                value = path + "#" + Encode(fragment);
            }
            // trailing white spaces must be url-encoded, see Bugzilla 53282
            if (value.Length > 0)
            {
                StringBuilder b = new StringBuilder();
                int idx = value.Length - 1;
                for (; idx >= 0; idx--)
                {
                    char c = value[idx];
                    if (char.IsWhiteSpace(c) || c == '\u00A0')
                    {
                        b.Append(c);
                    }
                    else
                    {
                        break;
                    }
                }
                if (b.Length > 0)
                {
                    char[] ca = b.ToString().ToCharArray();
                    Array.Reverse(ca);
                    value = value.Substring(0, idx + 1) + Encode(new string(ca));
                }
            }

            // MS Office can insert URIs with missing authority, e.g. "http://" or "javascript://"
            // append a forward slash to avoid parse exception
            if (missingAuthPattern.IsMatch(value))
            {
                value += "/";
            }
            return ParseUri(value, UriKind.RelativeOrAbsolute);  //unicode character is not allowed in Uri class before .NET4.5
        }

           /**
     * percent-encode white spaces and characters above 0x80.
     * <p>
     *   Examples:
     *   'Apache POI' --> 'Apache%20POI'
     *   'Apache\u0410POI' --> 'Apache%04%10POI'
     *
     * @param s the string to encode
     * @return  the encoded string
     */
    public static String Encode(String s) {
        int n = s.Length;
        if (n == 0) return s;

        byte[] bb = Encoding.UTF8.GetBytes(s);
        StringBuilder sb = new StringBuilder();
        foreach (byte b in bb)
        { 
            int b1 = (int)b & 0xff;
            if (IsUnsafe(b1)) {
                sb.Append('%');
                sb.Append(hexDigits[(b1 >> 4) & 0x0F]);
                sb.Append(hexDigits[(b1 >> 0) & 0x0F]);
            } else {
                sb.Append((char)b1);
            }           
        }
        return sb.ToString();
    }

    private static char[] hexDigits = {
        '0', '1', '2', '3', '4', '5', '6', '7',
        '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'
    };

    private static bool IsUnsafe(int ch)
    {
        return ch > 0x80 || char.IsWhiteSpace((char)ch) || ch == '\u00A0';
    }
        /**
         * Build a part name where the relationship should be stored ((ex
         * /word/document.xml -> /word/_rels/document.xml.rels)
         *
         * @param partName
         *            Source part Uri
         * @return the full path (as Uri) of the relation file
         * @throws InvalidOperationException
         *             Throws if the specified Uri is a relationshp part.
         */
        public static PackagePartName GetRelationshipPartName(
                PackagePartName partName)
        {
            if (partName == null)
                throw new ArgumentException("partName");

            if (PackagingUriHelper.PACKAGE_ROOT_URI.OriginalString == partName.URI
                    .OriginalString)
                return PackagingUriHelper.PACKAGE_RELATIONSHIPS_ROOT_PART_NAME;

            if (partName.IsRelationshipPartURI())
                throw new InvalidOperationException("Can't be a relationship part");

            String fullPath = partName.URI.OriginalString;
            String filename = GetFilename(partName.URI);
            fullPath = fullPath.Substring(0, fullPath.Length - filename.Length);
            fullPath = Combine(fullPath,
                    PackagingUriHelper.RELATIONSHIP_PART_SEGMENT_NAME);
            fullPath = Combine(fullPath, filename);
            fullPath = fullPath
                    + PackagingUriHelper.RELATIONSHIP_PART_EXTENSION_NAME;

            PackagePartName retPartName;
            try
            {
                retPartName = CreatePartName(fullPath);
            }
            catch (InvalidFormatException)
            {
                // Should never happen in production as all data are fixed but in
                // case of return null:
                return null;
            }
            return retPartName;
        }
    }
}