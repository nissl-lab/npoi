
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
using System.Collections;
using System.Globalization;

namespace NPOI.Util
{
    public abstract class POILogger
    {

        public const int DEBUG = 1;
        public const int INFO  = 3;
        public const int WARN  = 5;
        public const int ERROR = 7;
        public const int FATAL = 9;

        /**
         * package scope so it cannot be instantiated outside of the util
         * package. You need a POILogger? Go to the POILogFactory for one
         *
         */
        public POILogger()
        {}
        
        abstract public void Initialize(String cat);
        
        /**
         * Log a message
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 The object to Log.  This is converted to a string.
         */
        abstract public void Log(int level, Object obj1);
        
        /**
         * Log a message
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 The object to Log.  This is converted to a string.
         * @param exception An exception to be Logged
         */
        abstract public void Log(int level, Object obj1,
                        Exception exception);


        /**
         * Check if a Logger is enabled to Log at the specified level
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         */
        abstract public bool Check(int level);

        /*
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first object to place in the message
         * @param obj2 second object to place in the message
         */

       /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first object to place in the message
         * @param obj2 second object to place in the message
         */

        public virtual void Log(int level, Object obj1, Object obj2)
        {
            if (Check(level))
            {
                Log(level, new StringBuilder(32).Append(obj1).Append(obj2));
            }
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param obj3 third Object to place in the message
         */

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3)
        {
            

            if (Check(level))
            {
                Log(level,
                        new StringBuilder(48).Append(obj1).Append(obj2)
                            .Append(obj3));
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4)
        {
            

            if (Check(level))
            {
                Log(level,
                        new StringBuilder(64).Append(obj1).Append(obj2)
                            .Append(obj3).Append(obj4));
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5)
        {
            

            if (Check(level))
            {
                Log(level,
                        new StringBuilder(80).Append(obj1).Append(obj2)
                            .Append(obj3).Append(obj4).Append(obj5));
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6)
        {
            

            if (Check(level))
            {
                Log(level ,
                        new StringBuilder(96).Append(obj1).Append(obj2)
                            .Append(obj3).Append(obj4).Append(obj5).Append(obj6));
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Object obj7)
        {
            

            if (Check(level))
            {
                Log(level,
                        new StringBuilder(112).Append(obj1).Append(obj2)
                            .Append(obj3).Append(obj4).Append(obj5).Append(obj6)
                            .Append(obj7));
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Object obj7, Object obj8)
        {
            

            if (Check(level))
            {
                Log(level,
                        new StringBuilder(128).Append(obj1).Append(obj2)
                            .Append(obj3).Append(obj4).Append(obj5).Append(obj6)
                            .Append(obj7).Append(obj8));
            }
        }

        /**
         * Log an exception, without a message
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param exception An exception to be Logged
         */

        public virtual void Log(int level, Exception exception)
        {
            Log(level, null, exception);
        }

        /**
         * Log a message. Lazily appends Object parameters together.
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 first Object to place in the message
         * @param obj2 second Object to place in the message
         * @param exception An exception to be Logged
         */

        public virtual void Log(int level, Object obj1, Object obj2,
                        Exception exception)
        {
            

            if (Check(level))
            {
                Log(level, new StringBuilder(32).Append(obj1).Append(obj2),
                        exception);
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Exception exception)
        {
            

            if (Check(level))
            {
                Log(level, new StringBuilder(48).Append(obj1).Append(obj2)
                    .Append(obj3), exception);
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4,
                        Exception exception)
        {
            

            if (Check(level))
            {
                Log(level, new StringBuilder(64).Append(obj1).Append(obj2)
                    .Append(obj3).Append(obj4), exception);
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Exception exception)
        {
            

            if (Check(level))
            {
                Log(level, new StringBuilder(80).Append(obj1).Append(obj2)
                    .Append(obj3).Append(obj4).Append(obj5), exception);
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Exception exception)
        {
            

            if (Check(level))
            {
                Log(level , new StringBuilder(96).Append(obj1)
                    .Append(obj2).Append(obj3).Append(obj4).Append(obj5)
                    .Append(obj6), exception);
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Object obj7,
                        Exception exception)
        {
            

            if (Check(level))
            {
                Log(level, new StringBuilder(112).Append(obj1).Append(obj2)
                    .Append(obj3).Append(obj4).Append(obj5).Append(obj6)
                    .Append(obj7), exception);
            }
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

        public virtual void Log(int level, Object obj1, Object obj2,
                        Object obj3, Object obj4, Object obj5,
                        Object obj6, Object obj7, Object obj8,
                        Exception exception)
        {
            

            if (Check(level))
            {
                Log(level, new StringBuilder(128).Append(obj1).Append(obj2)
                    .Append(obj3).Append(obj4).Append(obj5).Append(obj6)
                    .Append(obj7).Append(obj8), exception);
            }
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

        public virtual void LogFormatted(int level, String message,
                                 Object obj1)
        {
            CommonLogFormatted(level, message, new Object[]
            {
                obj1
            });
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

        public virtual void LogFormatted(int level, String message,
                                 Object obj1, Object obj2)
        {
            CommonLogFormatted(level, message, new Object[]
            {
                obj1, obj2
            });
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

        public virtual void LogFormatted(int level, String message,
                                 Object obj1, Object obj2,
                                 Object obj3)
        {
            CommonLogFormatted(level, message, new Object[]
            {
                obj1, obj2, obj3
            });
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

        public virtual void LogFormatted(int level, String message,
                                 Object obj1, Object obj2,
                                 Object obj3, Object obj4)
        {
            CommonLogFormatted(level, message, new Object[]
            {
                obj1, obj2, obj3, obj4
            });
        }

        private void CommonLogFormatted(int level, String message,
                                        Object [] unflatParams)
        {
            

            if (Check(level))
            {
                object[] params1 = (object[])FlattenArrays(unflatParams);

                if (params1[ params1.Length - 1 ].GetType()==typeof(Exception))
                {
                    Log(level, string.Format(CultureInfo.InvariantCulture, message, params1),
                        ( Exception ) params1[ params1.Length - 1 ]);
                }
                else
                {
                    Log(level, string.Format(CultureInfo.InvariantCulture, message, params1));
                }
            }
        }

        /**
         * Flattens any contained objects. Only tranverses one level deep.
         */

        private Array FlattenArrays(Object [] objects)
        {
            ArrayList results = new ArrayList();

            for (int i = 0; i < objects.Length; i++)
            {
                results.AddRange(ObjectToObjectArray(objects[ i ]));
            }
            return results.ToArray();
        }

        private ArrayList ObjectToObjectArray(Object obj)
        {
            ArrayList results = new ArrayList();

            if (obj.GetType()==typeof(char[]))
            {
                byte[] array = ( byte [] ) obj;

                for (int j = 0; j < array.Length; j++)
                {
                    results.Add((byte)array[j]);
                }
            }
            if (obj.GetType()==typeof(char[]))
            {
                char[] array = ( char [] ) obj;

                for (int j = 0; j < array.Length; j++)
                {
                    results.Add((char)(array[ j ]));
                }
            }
            else if (obj.GetType()==typeof(short[]))
            {
                short[] array = ( short [] ) obj;

                for (int j = 0; j < array.Length; j++)
                {
                    results.Add((short)array[ j ]);
                }
            }
            else if (obj.GetType()==typeof(int[]))
            {
                int[] array = ( int [] ) obj;

                for (int j = 0; j < array.Length; j++)
                {
                    results.Add((int)array[ j ]);
                }
            }
            else if (obj.GetType()==typeof(long[]))
            {
                long[] array = ( long [] ) obj;

                for (int j = 0; j < array.Length; j++)
                {
                    results.Add((long)array[ j ]);
                }
            }
            else if (obj.GetType()==typeof(float []))
            {
                float[] array = ( float [] ) obj;

                for (int j = 0; j < array.Length; j++)
                {
                    results.Add((float)array[ j ]);
                }
            }
            else if (obj.GetType()==typeof(double []))
            {
                double[] array = ( double [] ) obj;

                for (int j = 0; j < array.Length; j++)
                {
                    results.Add((double)array[ j ]);
                }
            }
            else if (obj.GetType()==typeof(object[]))
            {
                object[] array = ( object[]) obj;

                for (int j = 0; j < array.Length; j++)
                {
                    results.Add(array[ j ]);
                }
            }
            else
            {
                results.Add(obj);
            }
            return results;
        }
                                     
    }
}
