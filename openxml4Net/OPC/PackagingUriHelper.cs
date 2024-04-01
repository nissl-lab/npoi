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
using System.Text;
using System.Text.RegularExpressions;
using NPOI.OpenXml4Net.Exceptions;
using System.IO;

namespace NPOI.OpenXml4Net.OPC
{
    /// <summary>
    /// Helper for part and pack Uri.
    /// </summary>
    /// <remarks>
    /// @author Julien Chable, CDubet, Kim Ung
    /// @version 0.1
    /// </remarks>

    public class PackagingUriHelper
    {

        /// <summary>
        /// Package root Uri.
        /// </summary>
        private static Uri packageRootUri;

        /// <summary>
        /// Extension name of a relationship part.
        /// </summary>
        public static String RELATIONSHIP_PART_EXTENSION_NAME;

        /// <summary>
        /// Segment name of a relationship part.
        /// </summary>
        public static String RELATIONSHIP_PART_SEGMENT_NAME;

        /// <summary>
        /// Segment name of the package properties folder.
        /// </summary>
        public static String PACKAGE_PROPERTIES_SEGMENT_NAME;

        /// <summary>
        /// Core package properties art name.
        /// </summary>
        public static String PACKAGE_CORE_PROPERTIES_NAME;

        /// <summary>
        /// Forward slash Uri separator.
        /// </summary>
        public static char FORWARD_SLASH_CHAR;

        /// <summary>
        /// Forward slash Uri separator.
        /// </summary>
        public static String FORWARD_SLASH_STRING;

        /// <summary>
        /// Package relationships part Uri
        /// </summary>
        public static Uri PACKAGE_RELATIONSHIPS_ROOT_URI;

        /// <summary>
        /// Package relationships part name.
        /// </summary>
        public static PackagePartName PACKAGE_RELATIONSHIPS_ROOT_PART_NAME;

        /// <summary>
        /// Core properties part Uri.
        /// </summary>
        public static Uri CORE_PROPERTIES_URI;

        /// <summary>
        /// Core properties partname.
        /// </summary>
        public static PackagePartName CORE_PROPERTIES_PART_NAME;

        /// <summary>
        /// Root package Uri.
        /// </summary>
        public static Uri PACKAGE_ROOT_URI;

        /// <summary>
        /// Root package part name.
        /// </summary>
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

            uriPACKAGE_ROOT_URI = ParseUri("/", UriKind.Relative);
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
            catch(InvalidFormatException)
            {
                // Should never happen in production as all data are fixed
            }
            PACKAGE_RELATIONSHIPS_ROOT_PART_NAME = tmpPACKAGE_RELATIONSHIPS_ROOT_PART_NAME;
            CORE_PROPERTIES_PART_NAME = tmpCORE_PROPERTIES_URI;
            PACKAGE_ROOT_PART_NAME = tmpPACKAGE_ROOT_PART_NAME;
        }
        private static Regex missingAuthPattern = new Regex("\\w+://$", RegexOptions.Compiled);
        /// <summary>
        /// Gets the Uri for the package root.
        /// </summary>
        /// <returns>Uri of the package root.</returns>
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
            if(IsMono)
            {
                if(kind == UriKind.Absolute)
                    throw new UriFormatException();
                if(kind == UriKind.RelativeOrAbsolute && s.StartsWith("/"))
                    kind = UriKind.Relative;
            }
            return new Uri(s, kind);
        }


        /// <summary>
        /// Know if the specified Uri is a relationship part name.
        /// </summary>
        /// <param name="partUri">
        /// Uri to check.
        /// </param>
        /// <returns><i>true</i> if the Uri <i>false</i>.</returns>
        public static bool IsRelationshipPartURI(Uri partUri)
        {
            if(partUri == null)
                throw new ArgumentException("partUri");

            return Regex.IsMatch(partUri.OriginalString,
                    ".*" + RELATIONSHIP_PART_SEGMENT_NAME + ".*"
                            + RELATIONSHIP_PART_EXTENSION_NAME + "$");
        }

        /// <summary>
        /// Get file name from the specified Uri.
        /// </summary>
        public static String GetFilename(Uri uri)
        {
            if(uri != null)
            {
                String path = uri.OriginalString;
                int len = path.Length;
                int num2 = len;
                while(--num2 >= 0)
                {
                    char ch1 = path[num2];
                    if(ch1 == PackagingUriHelper.FORWARD_SLASH_CHAR)
                        return path.Substring(num2 + 1);
                }
            }
            return "";
        }

        /// <summary>
        /// Get the file name without the trailing extension.
        /// </summary>
        public static String GetFilenameWithoutExtension(Uri uri)
        {
            String filename = GetFilename(uri);
            int dotIndex = filename.LastIndexOf(".", StringComparison.Ordinal);
            if(dotIndex == -1)
                return filename;
            return filename.Substring(0, dotIndex);
        }

        /// <summary>
        /// Get the directory path from the specified Uri.
        /// </summary>
        public static Uri GetPath(Uri uri)
        {
            if(uri != null)
            {
                String path = uri.OriginalString;
                int len = path.Length;
                int num2 = len;
                while(--num2 >= 0)
                {
                    char ch1 = path[num2];
                    if(ch1 == PackagingUriHelper.FORWARD_SLASH_CHAR)
                    {
                        try
                        {
                            return ParseUri(path.Substring(0, num2), UriKind.Absolute);
                        }
                        catch(UriFormatException)
                        {
                            return null;
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Combine two URIs.
        /// </summary>
        /// <param name="prefix">the prefix Uri</param>
        /// <param name="suffix">the suffix Uri</param>
        /// 
        /// <returns>the Combined Uri</returns>
        public static Uri Combine(Uri prefix, Uri suffix)
        {
            Uri retUri = null;
            try
            {
                retUri = ParseUri(Combine(prefix.OriginalString, suffix.OriginalString), UriKind.Absolute);
            }
            catch(UriFormatException)
            {
                throw new ArgumentException(
                        "Prefix and suffix can't be Combine !");
            }
            return retUri;
        }

        /// <summary>
        /// Combine a string Uri with a prefix and a suffix.
        /// </summary>
        public static String Combine(String prefix, String suffix)
        {
            if(!prefix.EndsWith("" + FORWARD_SLASH_CHAR)
                    && !suffix.StartsWith("" + FORWARD_SLASH_CHAR))
                return prefix + FORWARD_SLASH_CHAR + suffix;
            else if((!prefix.EndsWith("" + FORWARD_SLASH_CHAR)
                    && suffix.StartsWith("" + FORWARD_SLASH_CHAR) || (prefix
                    .EndsWith("" + FORWARD_SLASH_CHAR) && !suffix.StartsWith(""
                    + FORWARD_SLASH_CHAR))))
                return prefix + suffix;
            else
                return "";
        }

        /// <summary>
        /// Fully relativize the source part Uri against the target part Uri.
        /// </summary>
        /// <param name="sourceURI">
        /// The source part Uri.
        /// </param>
        /// <param name="targetURI">
        /// The target part Uri.
        /// </param>
        /// <returns>A fully relativize part name Uri ('word/media/image1.gif',
        /// '/word/document.xml' => 'media/image1.gif') else
        /// <c>null</c>.
        /// </returns>
        public static Uri RelativizeUri(Uri sourceURI, Uri targetURI, bool msCompatible)
        {
            StringBuilder retVal = new StringBuilder();
            String[] segmentsSource = sourceURI.ToString().Split(new char[] { '/' });
            String[] segmentsTarget = targetURI.ToString().Split(new char[] { '/' });

            // If the source Uri is empty
            if(segmentsSource.Length == 0)
            {
                throw new ArgumentException(
                        "Can't relativize an empty source Uri !");
            }

            // If target Uri is empty
            if(segmentsTarget.Length == 0)
            {
                throw new ArgumentException(
                        "Can't relativize an empty target Uri !");
            }

            // If the source is the root, then the relativized
            //  form must actually be an absolute Uri
            if(sourceURI.ToString().Equals("/"))
            {
                String path = targetURI.ToString();
                if(msCompatible && path.Length > 0 && path[0] == '/')
                {
                    try
                    {
                        targetURI = ParseUri(path.Substring(1), UriKind.RelativeOrAbsolute);
                    }
                    catch
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
            for(int i = 0; i < segmentsSource.Length && i < segmentsTarget.Length; i++)
            {
                if(segmentsSource[i].Equals(segmentsTarget[i]))
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
            if((segmentsTheSame == 0 || segmentsTheSame == 1) &&
                    segmentsSource[0].Equals("") && segmentsTarget[0].Equals(""))
            {
                for(int i = 0; i < segmentsSource.Length - 2; i++)
                {
                    retVal.Append("../");
                }
                for(int i = 0; i < segmentsTarget.Length; i++)
                {
                    if(segmentsTarget[i].Equals(""))
                        continue;
                    retVal.Append(segmentsTarget[i]);
                    if(i != segmentsTarget.Length - 1)
                        retVal.Append("/");
                }

                try
                {
                    return ParseUri(retVal.ToString(), UriKind.RelativeOrAbsolute);
                }
                catch
                {
                    //System.err.println(e);
                    return null;
                }
            }

            // Special case for where the two are the same
            if(segmentsTheSame == segmentsSource.Length
                    && segmentsTheSame == segmentsTarget.Length)
            {
                if(sourceURI.Equals(targetURI))
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
                if(segmentsTheSame == 1)
                {
                    retVal.Append("/");
                }
                else
                {
                    for(int j = segmentsTheSame; j < segmentsSource.Length - 1; j++)
                    {
                        retVal.Append("../");
                    }
                }

                // Now go from here on down
                for(int j = segmentsTheSame; j < segmentsTarget.Length; j++)
                {
                    if(retVal.Length > 0
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
            catch
            {
                //System.err.println(e);
                return null;
            }
        }


        /// <summary>
        /// Fully relativize the source part URI against the target part URI.
        /// </summary>
        /// <param name="sourceURI">
        /// The source part URI.
        /// </param>
        /// <param name="targetURI">
        /// The target part URI.
        /// </param>
        /// <returns>A fully relativize part name URI ('word/media/image1.gif',
        /// '/word/document.xml' => 'media/image1.gif') else
        /// <c>null</c>.
        /// </returns>
        public static Uri RelativizeUri(Uri sourceURI, Uri targetURI)
        {
            return RelativizeUri(sourceURI, targetURI, false);
        }

        /// <summary>
        /// Resolve a source uri against a target.
        /// </summary>
        /// <param name="sourcePartUri">
        /// The source Uri.
        /// </param>
        /// <param name="targetUri">
        /// The target Uri.
        /// </param>
        /// <returns>The resolved Uri.</returns>
        public static Uri ResolvePartUri(Uri sourcePartUri, Uri targetUri)
        {
            if(sourcePartUri == null || sourcePartUri.IsAbsoluteUri)
            {
                throw new ArgumentException("sourcePartUri invalid - "
                        + sourcePartUri);
            }

            if(targetUri == null || targetUri.IsAbsoluteUri)
            {
                throw new ArgumentException("targetUri invalid - "
                        + targetUri);
            }
            string path;
            if(sourcePartUri.OriginalString == "/")
                path = "/";
            else
                path = Path.GetDirectoryName(sourcePartUri.OriginalString).Replace("\\", "/");

            string targetPath = targetUri.OriginalString;
            if(targetPath.StartsWith("#"))
            {
                path += "/" + Path.GetFileName(sourcePartUri.OriginalString) + targetPath;
            }
            else if(targetPath.StartsWith("../"))
            {
                string[] segments = path.Split(new char[] { '/' });

                int segmentEnd = segments.Length - 1;
                while(targetPath.StartsWith("../"))
                {
                    targetPath = targetPath.Substring(3);
                    segmentEnd -= 1;
                }
                path = "/";

                for(int i = 0; i <= segmentEnd; i++)
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

        /// <summary>
        /// Get Uri from a string path.
        /// </summary>
        public static Uri GetURIFromPath(String path)
        {
            Uri retUri = null;
            try
            {
                retUri = ParseUri(path, UriKind.RelativeOrAbsolute);
            }
            catch(UriFormatException)
            {
                throw new ArgumentException("path");
            }
            return retUri;
        }

        /// <summary>
        /// Get the source part Uri from a specified relationships part.
        /// </summary>
        /// <param name="relationshipPartUri">
        /// The relationship part use to retrieve the source part.
        /// </param>
        /// <returns>The source part Uri from the specified relationships part.</returns>
        public static Uri GetSourcePartUriFromRelationshipPartUri(
                Uri relationshipPartUri)
        {
            if(relationshipPartUri == null)
                throw new ArgumentException(
                        "Must not be null");

            if(!IsRelationshipPartURI(relationshipPartUri))
                throw new ArgumentException(
                        "Must be a relationship part");

            if(Uri.Compare(relationshipPartUri, PACKAGE_RELATIONSHIPS_ROOT_URI, UriComponents.AbsoluteUri, UriFormat.SafeUnescaped, StringComparison.InvariantCultureIgnoreCase) == 0)
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

        /// <summary>
        /// Create an OPC compliant part name by throwing an exception if the Uri is
        /// not valid.
        /// </summary>
        /// <param name="partUri">
        /// The part name Uri to validate.
        /// </param>
        /// <returns>A valid part name object, else <c>null</c>.</returns>
        /// <exception cref="InvalidFormatException">
        /// Throws if the specified Uri is not OPC compliant.
        /// </exception>
        public static PackagePartName CreatePartName(Uri partUri)
        {
            if(partUri == null)
                throw new ArgumentException("partName");

            return new PackagePartName(partUri, true);
        }

        /// <summary>
        /// Create an OPC compliant part name.
        /// </summary>
        /// <param name="partName">
        /// The part name to validate.
        /// </param>
        /// <returns>The correspondant part name if valid, else <c>null</c>.</returns>
        /// <exception cref="InvalidFormatException">
        /// Throws if the specified part name is not OPC compliant.
        /// </exception>
        /// @see #CreatePartName(Uri)
        public static PackagePartName CreatePartName(String partName)
        {
            Uri partNameURI;
            try
            {
                partName = partName.Replace("\\", "/");  //tolerate backslash - poi test49609
                partNameURI = ParseUri(partName, UriKind.Relative);
            }
            catch(UriFormatException e)
            {
                throw new InvalidFormatException(e.Message);
            }
            return CreatePartName(partNameURI);
        }

        /// <summary>
        /// Create an OPC compliant part name by resolving it using a base part.
        /// </summary>
        /// <param name="partName">
        /// The part name to validate.
        /// </param>
        /// <param name="relativePart">
        /// The relative base part.
        /// </param>
        /// <returns>The correspondant part name if valid, else <c>null</c>.</returns>
        /// <exception cref="InvalidFormatException">
        /// Throws if the specified part name is not OPC compliant.
        /// </exception>
        /// @see #CreatePartName(Uri)
        public static PackagePartName CreatePartName(String partName,
                PackagePart relativePart)
        {
            Uri newPartNameURI;
            try
            {
                newPartNameURI = ResolvePartUri(
                        relativePart.PartName.URI, ParseUri(partName, UriKind.RelativeOrAbsolute));
            }
            catch(UriFormatException e)
            {
                throw new InvalidFormatException(e.Message);
            }
            return CreatePartName(newPartNameURI);
        }

        /// <summary>
        /// Create an OPC compliant part name by resolving it using a base part.
        /// </summary>
        /// <param name="partName">
        /// The part name Uri to validate.
        /// </param>
        /// <param name="relativePart">
        /// The relative base part.
        /// </param>
        /// <returns>The correspondant part name if valid, else <c>null</c>.</returns>
        /// <exception cref="InvalidFormatException">
        /// Throws if the specified part name is not OPC compliant.
        /// </exception>
        /// @see #CreatePartName(Uri)
        public static PackagePartName CreatePartName(Uri partName,
                PackagePart relativePart)
        {
            Uri newPartNameURI = ResolvePartUri(
                    relativePart.PartName.URI, partName);
            return CreatePartName(newPartNameURI);
        }

        /// <summary>
        /// <para>
        /// Validate a part Uri by returning a bool.
        /// ([M1.1],[M1.3],[M1.4],[M1.5],[M1.6])
        /// </para>
        /// <para>
        /// (OPC Specifications 8.1.1 Part names) :
        /// </para>
        /// <para>
        /// Part Name Syntax
        /// </para>
        /// <para>
        /// The part name grammar is defined as follows:
        /// </para>
        /// <para>
        /// <i>part_name = 1*( "/" segment )
        /// </para>
        /// <para>
        /// segment = 1*( pchar )</i>
        /// </para>
        /// <para>
        /// (pchar is defined in RFC 3986)
        /// </para>
        /// </summary>
        /// <param name="partUri">
        /// The Uri to validate.
        /// </param>
        /// <returns><b>true</b> if the Uri is valid to the OPC Specifications, else
        /// <b>false</b>
        /// </returns>
        /// 
        /// @see #CreatePartName(Uri)
        public static bool IsValidPartName(Uri partUri)
        {
            if(partUri == null)
                throw new ArgumentException("partUri");

            try
            {
                CreatePartName(partUri);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Decode a Uri by converting all percent encoded character into a String
        /// character.
        /// </summary>
        /// <param name="uri">
        /// The Uri to decode.
        /// </param>
        /// <returns>The specified Uri in a String with converted percent encoded
        /// characters.
        /// </returns>
        public static String DecodeURI(Uri uri)
        {
            StringBuilder retVal = new StringBuilder();
            String uriStr = uri.OriginalString;
            char c;
            int length = uriStr.Length;
            for(int i = 0; i < length; ++i)
            {
                c = uriStr[i];
                if(c == '%')
                {
                    // We certainly found an encoded character, check for length
                    // now ( '%' HEXDIGIT HEXDIGIT)
                    if((length - i) < 2)
                    {
                        throw new ArgumentException("The uri " + uriStr
                                + " contain invalid encoded character !");
                    }

                    // Decode the encoded character
                    char decodedChar = (char)Convert.ToInt32(uriStr.Substring(
                            i + 1, 2), 16);
                    retVal.Append(decodedChar);
                    i += 2;
                    continue;
                }
                retVal.Append(c);
            }
            return retVal.ToString();
        }
        /// <summary>
        /// <para>
        /// Convert a string to <see cref="java.net.URI" />
        /// </para>
        /// <para>
        /// If  part name is not a valid URI, it is resolved as follows:
        /// </para>
        /// <para>
        /// 1. Percent-encode each open bracket ([) and close bracket (]).</li>
        /// 2. Percent-encode each percent (%) character that is not followed by a hexadecimal notation of an octet value.</li>
        /// 3. Un-percent-encode each percent-encoded unreserved character.
        /// 4. Un-percent-encode each forward slash (/) and back slash (\).
        /// 5. Convert all back slashes to forward slashes.
        /// 6. If present in a segment containing non-dot (?.?) characters, remove trailing dot (?.?) characters from each segment.
        /// 7. Replace each occurrence of multiple consecutive forward slashes (/) with a single forward slash.
        /// 8. If a single trailing forward slash (/) is present, remove that trailing forward slash.
        /// 9. Remove complete segments that consist of three or more dots.
        /// 10. Resolve the relative reference against the base URI of the part holding the Unicode string, as it is defined
        /// in ?5.2 of RFC 3986. The path component of the resulting absolute URI is the part name.
        /// </para>
        /// </summary>
        /// <param name="value">  the string to be parsed into a URI</param>
        /// <returns>the resolved part name that should be OK to construct a URI</returns>
        /// 
        /// TODO YK: for now this method does only (5). Finish the rest.
        public static Uri ToUri(String value)
        {
            //5. Convert all back slashes to forward slashes
            if(value.IndexOf("\\") != -1)
            {
                value = value.Replace('\\', '/');
            }

            // URI fragemnts (those starting with '#') are not encoded
            // and may contain white spaces and raw unicode characters
            int fragmentIdx = value.IndexOf('#');
            if(fragmentIdx != -1)
            {
                String path = value.Substring(0, fragmentIdx);
                String fragment = value.Substring(fragmentIdx + 1);

                value = path + "#" + Encode(fragment);
            }
            // trailing white spaces must be url-encoded, see Bugzilla 53282
            if(value.Length > 0)
            {
                StringBuilder b = new StringBuilder();
                int idx = value.Length - 1;
                for(; idx >= 0; idx--)
                {
                    char c = value[idx];
                    if(char.IsWhiteSpace(c) || c == '\u00A0')
                    {
                        b.Append(c);
                    }
                    else
                    {
                        break;
                    }
                }
                if(b.Length > 0)
                {
                    char[] ca = b.ToString().ToCharArray();
                    Array.Reverse(ca);
                    value = value.Substring(0, idx + 1) + Encode(new string(ca));
                }
            }

            // MS Office can insert URIs with missing authority, e.g. "http://" or "javascript://"
            // append a forward slash to avoid parse exception
            if(missingAuthPattern.IsMatch(value))
            {
                value += "/";
            }
            return ParseUri(value, UriKind.RelativeOrAbsolute);  //unicode character is not allowed in Uri class before .NET4.5
        }

        /// <summary>
        /// <para>
        /// percent-encode white spaces and characters above 0x80.
        /// </para>
        /// <para>
        ///   Examples:
        ///   'Apache POI' --> 'Apache%20POI'
        ///   'Apache\u0410POI' --> 'Apache%04%10POI'
        /// </para>
        /// </summary>
        /// <param name="s">the string to encode</param>
        /// <returns>the encoded string</returns>
        public static String Encode(String s)
        {
            int n = s.Length;
            if(n == 0)
                return s;

            byte[] bb = Encoding.UTF8.GetBytes(s);
            StringBuilder sb = new StringBuilder();
            foreach(byte b in bb)
            {
                int b1 = (int)b & 0xff;
                if(IsUnsafe(b1))
                {
                    sb.Append('%');
                    sb.Append(hexDigits[(b1 >> 4) & 0x0F]);
                    sb.Append(hexDigits[(b1 >> 0) & 0x0F]);
                }
                else
                {
                    sb.Append((char) b1);
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
            return ch > 0x80 || char.IsWhiteSpace((char) ch);
        }
        /// <summary>
        /// Build a part name where the relationship should be stored ((ex
        /// /word/document.xml -> /word/_rels/document.xml.rels)
        /// </summary>
        /// <param name="partName">
        /// Source part Uri
        /// </param>
        /// <returns>the full path (as Uri) of the relation file</returns>
        /// <exception cref="InvalidOperationException">
        /// Throws if the specified Uri is a relationshp part.
        /// </exception>
        public static PackagePartName GetRelationshipPartName(
                PackagePartName partName)
        {
            if(partName == null)
                throw new ArgumentException("partName");

            if(PackagingUriHelper.PACKAGE_ROOT_URI.OriginalString == partName.URI
                    .OriginalString)
                return PackagingUriHelper.PACKAGE_RELATIONSHIPS_ROOT_PART_NAME;

            if(partName.IsRelationshipPartURI())
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
            catch(InvalidFormatException)
            {
                // Should never happen in production as all data are fixed but in
                // case of return null:
                return null;
            }
            return retPartName;
        }
    }
}