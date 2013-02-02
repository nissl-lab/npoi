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

/* ================================================================
 * About NPOI
 * Author: Tony Qu 
 * Author's email: tonyqus (at) gmail.com 
 * Author's Blog: tonyqus.wordpress.com.cn (wp.tonyqus.cn)
 * HomePage: http://www.codeplex.com/npoi
 * Contributors:
 * 
 * ==============================================================*/

using System;
using System.Text;
using System.IO;
using System.Collections;

namespace NPOI.Util
{
    public class HexRead
    {
        /// <summary>
        /// This method reads hex data from a filename and returns a byte array.
        /// The file may contain line comments that are preceeded with a # symbol.
        /// </summary>
        /// <param name="filename">The filename to read</param>
        /// <returns>The bytes read from the file.</returns>
        /// <exception cref="IOException">If there was a problem while reading the file.</exception>
        public static byte[] ReadData( String filename )
        {
            FileStream stream = new FileStream(filename,FileMode.Open,FileAccess.Read);
            try
            {
                return ReadData( stream, -1 );
            }
            finally
            {
                stream.Close();
            }
        }

        /// <summary>
        /// Same as ReadData(String) except that this method allows you to specify sections within
        /// a file.  Sections are referenced using section headers in the form:
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="section">The section.</param>
        /// <returns></returns>
        public static byte[] ReadData(Stream stream, String section )
        {
        	
            try
            {
                StringBuilder sectionText = new StringBuilder();
                bool inSection = false;
                int c = stream.ReadByte();
                while ( c != -1 )
                {
                    switch ( c )
                    {
                        case '[':
                            inSection = true;
                            break;
                        case '\n':
                        case '\r':
                            inSection = false;
                            sectionText = new StringBuilder();
                            break;
                        case ']':
                            inSection = false;
                            if (sectionText.ToString().Equals(section))
                            {
                                return ReadData(stream, '[');
                            }
                            sectionText = new StringBuilder();
                            break;
                        default:
                            if ( inSection ) sectionText.Append( (char) c );
                            break;
                    }
                    c = stream.ReadByte();
                }
            }
            finally
            {
                stream.Close();
            }
            throw new IOException( "Section '" + section + "' not found" );
        }
        /// <summary>
        /// Reads the data.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <param name="section">The section.</param>
        /// <returns></returns>
        public static byte[] ReadData( String filename, String section )
        {
            using (FileStream stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                return ReadData(stream, section);
            }
        }

        /// <summary>
        /// Reads the data.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="eofChar">The EOF char.</param>
        /// <returns></returns>
        public static byte[] ReadData( Stream stream, int eofChar )
        {
            int characterCount = 0;
            byte b = (byte) 0;
            ArrayList bytes = new ArrayList();
            bool done = false;
            while ( !done )
            {
                int count = stream.ReadByte();
                char baseChar = 'a';
                if ( count == eofChar ) break;
                switch ( count )
                {
                    case '#':
                        ReadToEOL( stream );
                        break;
                    case '0': case '1': case '2': case '3': case '4': case '5':
                    case '6': case '7': case '8': case '9':
                        b <<= 4;
                        b += (byte) ( count - '0' );
                        characterCount++;
                        if ( characterCount == 2 )
                        {
                            bytes.Add( (byte)b );
                            characterCount = 0;
                            b = (byte) 0;
                        }
                        break;
                    case 'A':
                    case 'B':
                    case 'C':
                    case 'D':
                    case 'E':
                    case 'F':
                        baseChar = 'A';
                        b <<= 4;
                        b += (byte)(count + 10 - baseChar);
                        characterCount++;
                        if (characterCount == 2)
                        {
                            bytes.Add((byte)b);
                            characterCount = 0;
                            b = (byte)0;
                        }
                        break;
                    case 'a':
                    case 'b':
                    case 'c':
                    case 'd':
                    case 'e':
                    case 'f':
                        b <<= 4;
                        b += (byte) ( count + 10 - baseChar );
                        characterCount++;
                        if ( characterCount == 2 )
                        {
                            bytes.Add( (byte) b );
                            characterCount = 0;
                            b = (byte) 0;
                        }
                        break;
                    case -1:
                        done = true;
                        break;
                    default :
                        break;
                }
            }
            byte[] polished = (byte[]) bytes.ToArray(typeof(byte) );
            //byte[] rval = new byte[polished.Length];
            //for ( int j = 0; j < polished.Length; j++ )
            //{
            //    rval[j] = polished[j];
            //}
            return polished;
        }

        /// <summary>
        /// Reads from string.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        public static byte[] ReadFromString(String data)
        {
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(data)))
            {
                return ReadData(ms, -1);
            }
        }

        /// <summary>
        /// Reads to EOL.
        /// </summary>
        /// <param name="stream">The stream.</param>
        static private void ReadToEOL( Stream stream )
        {
            int c = stream.ReadByte();
            while ( c != -1 && c != '\n' && c != '\r' )
            {
                c = stream.ReadByte();
            }
        }
    }
}
