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
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace NPOI.POIFS.Crypt
{
    public class Base64  //Convert Jave Base64 Leon
    {
        private static char[] base64Code = new char[]
        {'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T',

　      　'U','V','W','X','Y','Z','a','b','c','d','e','f','g','h','i','j','k','l','m','n',

           'o','p','q','r','s','t','u','v','w','x','y','z','0','1','2','3','4','5','6','7',

　　       '8','9','+','/','='
        };

        private static string base64Str = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";

        public static string Base64Code(string message)
        {
            byte empty = (byte)0;
            ArrayList byteMessage = new ArrayList(Encoding.Default.GetBytes(message));

            StringBuilder outMessage;
            int messageLen = byteMessage.Count;

            int page = messageLen / 3;
            int use = 0;

            if ((use = messageLen % 3) > 0)
            {
                for (int i = 0; i < 3 - use; i++)
                {
                    byteMessage.Add(empty);
                }

                page++;
            }

            outMessage = new StringBuilder(page * 4);
            for (int i = 0; i < page; i++)
            {
                byte[] inStr = new byte[3];
                inStr[0] = (byte)byteMessage[i * 3];
                inStr[1] = (byte)byteMessage[i * 3 + 1];
                inStr[2] = (byte)byteMessage[i * 3 + 2];

                int[] outStr = new int[4];

                outStr[0] = inStr[0] >> 2;
                outStr[1] = ((inStr[0] & 0x03) << 4) ^ (inStr[1] >> 4);
                if (!inStr.Equals(empty))
                    outStr[2] = ((inStr[1] & 0x0f) << 2) ^ (inStr[2] >> 6);
                else
                    outStr[2] = 64;

                if (!inStr[2].Equals(empty))
                {
                    outStr[3] = inStr[2] & 0x3f;
                }
                else
                {
                    outStr[3] = 64;
                }

                outMessage.Append(base64Code[outStr[0]]);
                outMessage.Append(base64Code[outStr[1]]);
                outMessage.Append(base64Code[outStr[2]]);
                outMessage.Append(base64Code[outStr[3]]);
            }

            return outMessage.ToString();
        }

        public static byte[] DecodeBase64(string message)
        {
            if ((message.Length % 4) != 0)
            {
                throw new ArgumentException("Please check the correct base64 code");
            }
            if (!Regex.IsMatch(message, "^[A-Z0-9/+=]*$", RegexOptions.IgnoreCase))
            {
                throw new ArgumentException("Please check the correct base64 code");
            }

            int page = message.Length / 4;
            ArrayList outMessage = new ArrayList(page * 3);
            char[] chMessage = message.ToCharArray();

            for (int i = 0; i < page; i++)
            {
                byte[] inStr = new byte[4];
                inStr[0] = (byte)base64Str.IndexOf(message[i * 4]);
                inStr[1] = (byte)base64Str.IndexOf(message[i * 4 + 1]);
                inStr[2] = (byte)base64Str.IndexOf(message[i * 4 + 2]);
                inStr[3] = (byte)base64Str.IndexOf(message[i * 4 + 3]);

                byte[] outStr = new byte[3];
                outStr[0] = (byte)((inStr[0] << 2) ^ ((inStr[1] & 0x30) >> 4));

                if (inStr[2] != 64)
                {
                    outStr[1] = (byte)((inStr[1] << 4) ^ ((inStr[2] & 0x30) >> 2));
                }
                else
                {
                    outStr[1] = 0;
                }

                if (inStr[3] != 64)
                {
                    outStr[2] = (byte)((inStr[1] << 6) ^ inStr[3]);
                }
                else
                {
                    outStr[2] = 0;
                }

                outMessage.Add(outStr[0]);
                if (outStr[1] != 0)
                {
                    outMessage.Add(outStr[1]);
                }

                if (outStr[2] != 0)
                {
                    outMessage.Add(outStr[2]);
                }
            }
            byte[] outByte = (byte[])outMessage.ToArray(typeof(System.Byte));
            //return Encoding.Default.GetString(outByte);
            return outByte;
        }
    }
}











































