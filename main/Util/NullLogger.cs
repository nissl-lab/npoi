        
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

namespace NPOI.Util
{
    /// <summary>
    /// A Logger class that strives to make it as easy as possible for
    /// developers to write Log calls, while simultaneously making those
    /// calls as cheap as possible by performing lazy evaluation of the Log
    /// message.
    /// @author Marc Johnson (mjohnson at apache dot org)
    /// @author Glen Stampoultzis (glens at apache.org)
    /// @author Nicola Ken Barozzi (nicolaken at apache.org)
    /// </summary>
    public class NullLogger:POILogger
    {
        public override void Initialize(String cat)
        {
           //do nothing
        }
        
        /**
         * Log a message
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 The object to Log.
         */

        public override void Log(int level, Object obj1)
        {
            //do nothing
        }

        /**
         * Check if a Logger is enabled to Log at the specified level
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         */

        public override bool Check(int level)
        {
           return false;
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first object to place in the message
         * @param obj2 second object to place in the message
         */

        public override void Log(int level, Object obj1, Object obj2)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third Object to place in the message
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third Object to place in the message
         * @param obj4 fourth Object to place in the message
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third Object to place in the message
         * @param obj4 fourth Object to place in the message
         * @param obj5 fifth Object to place in the message
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third Object to place in the message
         * @param obj4 fourth Object to place in the message
         * @param obj5 fifth Object to place in the message
         * @param obj6 sixth Object to place in the message
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third Object to place in the message
         * @param obj4 fourth Object to place in the message
         * @param obj5 fifth Object to place in the message
         * @param obj6 sixth Object to place in the message
         * @param obj7 seventh Object to place in the message
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Object obj7)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third Object to place in the message
         * @param obj4 fourth Object to place in the message
         * @param obj5 fifth Object to place in the message
         * @param obj6 sixth Object to place in the message
         * @param obj7 seventh Object to place in the message
         * @param obj8 eighth Object to place in the message
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Object obj7, Object obj8)
        {
           //do nothing
        }

        /**
         * Log a message
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 The object to Log.  This is converted to a string.
         * @param exception An exception to be Logged
         */

        public override void Log(int level, Object obj1,
                        Exception exception)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param exception An exception to be Logged
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Exception exception)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third object to place in the message
         * @param exception An error message to be Logged
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Exception exception)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third object to place in the message
         * @param obj4 fourth object to place in the message
         * @param exception An exception to be Logged
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4,
                        Exception exception)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third object to place in the message
         * @param obj4 fourth object to place in the message
         * @param obj5 fifth object to place in the message
         * @param exception An exception to be Logged
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Exception exception)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third object to place in the message
         * @param obj4 fourth object to place in the message
         * @param obj5 fifth object to place in the message
         * @param obj6 sixth object to place in the message
         * @param exception An exception to be Logged
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Exception exception)
        {
           //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third object to place in the message
         * @param obj4 fourth object to place in the message
         * @param obj5 fifth object to place in the message
         * @param obj6 sixth object to place in the message
         * @param obj7 seventh object to place in the message
         * @param exception An exception to be Logged
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Object obj7,
                        Exception exception)
        {
          //do nothing
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third object to place in the message
         * @param obj4 fourth object to place in the message
         * @param obj5 fifth object to place in the message
         * @param obj6 sixth object to place in the message
         * @param obj7 seventh object to place in the message
         * @param obj8 eighth object to place in the message
         * @param exception An exception to be Logged
         */

        public override void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Object obj7, Object obj8,
                        Exception exception)
        {
           //do nothing
        }

        /**
         * Logs a formated message. The message itself may contain %
         * characters as place holders. This routine will attempt to match
         * the placeholder by looking at the type of parameter passed to
         * obj1.
         *
         * If the parameter is an array, it traverses the array first and
         * matches parameters sequentially against the array items.
         * Otherwise the parameters after <c>message</c> are matched
         * in order.
         *
         * If the place holder matches against a number it is printed as a
         * whole number. This can be overridden by specifying a precision
         * in the form %n.m where n is the padding for the whole part and
         * m is the number of decimal places to display. n can be excluded
         * if desired. n and m may not be more than 9.
         *
         * If the last parameter (after flattening) is a Exception it is
         * Logged specially.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param message The message to Log.
         * @param obj1 The first object to match against.
         */

        public override void LogFormatted(int level, String message,
                                 Object obj1)
        {
           //do nothing
        }

        /**
         * Logs a formated message. The message itself may contain %
         * characters as place holders. This routine will attempt to match
         * the placeholder by looking at the type of parameter passed to
         * obj1.
         *
         * If the parameter is an array, it traverses the array first and
         * matches parameters sequentially against the array items.
         * Otherwise the parameters after <c>message</c> are matched
         * in order.
         *
         * If the place holder matches against a number it is printed as a
         * whole number. This can be overridden by specifying a precision
         * in the form %n.m where n is the padding for the whole part and
         * m is the number of decimal places to display. n can be excluded
         * if desired. n and m may not be more than 9.
         *
         * If the last parameter (after flattening) is a Exception it is
         * Logged specially.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param message The message to Log.
         * @param obj1 The first object to match against.
         * @param obj2 The second object to match against.
         */

        public override void LogFormatted(int level, String message,
                                 Object obj1, Object obj2)
        {
           //do nothing
        }

        /**
         * Logs a formated message. The message itself may contain %
         * characters as place holders. This routine will attempt to match
         * the placeholder by looking at the type of parameter passed to
         * obj1.
         *
         * If the parameter is an array, it traverses the array first and
         * matches parameters sequentially against the array items.
         * Otherwise the parameters after <c>message</c> are matched
         * in order.
         *
         * If the place holder matches against a number it is printed as a
         * whole number. This can be overridden by specifying a precision
         * in the form %n.m where n is the padding for the whole part and
         * m is the number of decimal places to display. n can be excluded
         * if desired. n and m may not be more than 9.
         *
         * If the last parameter (after flattening) is a Exception it is
         * Logged specially.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param message The message to Log.
         * @param obj1 The first object to match against.
         * @param obj2 The second object to match against.
         * @param obj3 The third object to match against.
         */

        public override void LogFormatted(int level, String message,
                                 Object obj1, Object obj2,
                                 Object obj3)
        {
           //do nothing
        }

        /**
         * Logs a formated message. The message itself may contain %
         * characters as place holders. This routine will attempt to match
         * the placeholder by looking at the type of parameter passed to
         * obj1.
         *
         * If the parameter is an array, it traverses the array first and
         * matches parameters sequentially against the array items.
         * Otherwise the parameters after <c>message</c> are matched
         * in order.
         *
         * If the place holder matches against a number it is printed as a
         * whole number. This can be overridden by specifying a precision
         * in the form %n.m where n is the padding for the whole part and
         * m is the number of decimal places to display. n can be excluded
         * if desired. n and m may not be more than 9.
         *
         * If the last parameter (after flattening) is a Exception it is
         * Logged specially.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param message The message to Log.
         * @param obj1 The first object to match against.
         * @param obj2 The second object to match against.
         * @param obj3 The third object to match against.
         * @param obj4 The forth object to match against.
         */

        public override void LogFormatted(int level, String message,
                                 Object obj1, Object obj2,
                                 Object obj3, Object obj4)
        {
           //do nothing
        }
    }
}
