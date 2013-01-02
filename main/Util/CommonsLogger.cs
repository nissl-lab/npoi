
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
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


namespace NPOI.Util
{
    using System;

    /**
     * A logger class that strives to make it as easy as possible for
     * developers to write log calls, while simultaneously making those
     * calls as cheap as possible by performing lazy Evaluation of the log
     * message.<p>
     *
     * @author Marc Johnson (mjohnson at apache dot org)
     * @author Glen Stampoultzis (glens at apache.org)
     * @author Nicola Ken Barozzi (nicolaken at apache.org)
     */

    public class CommonsLogger : POILogger
    {

        private static LogFactory _creator = LogFactory.Factory;
        private Log log = null;


        public void Initialize(String cat)
        {
            this.log = _creator.GetInstance(cat);
        }

        /**
         * Log a message
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 The object to log.
         */
        public void log(int level, Object obj1)
        {
            if (level == FATAL)
            {
                if (log.IsFatalEnabled())
                {
                    log.Fatal(obj1);
                }
            }
            else if (level == ERROR)
            {
                if (log.IsErrorEnabled())
                {
                    log.Error(obj1);
                }
            }
            else if (level == WARN)
            {
                if (log.IsWarnEnabled())
                {
                    log.Warn(obj1);
                }
            }
            else if (level == INFO)
            {
                if (log.IsInfoEnabled())
                {
                    log.Info(obj1);
                }
            }
            else if (level == DEBUG)
            {
                if (log.IsDebugEnabled())
                {
                    log.Debug(obj1);
                }
            }
            else
            {
                if (log.IsTraceEnabled())
                {
                    log.Trace(obj1);
                }
            }
        }

        /**
         * Log a message
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         * @param obj1 The object to log.  This is Converted to a string.
         * @param exception An exception to be logged
         */
        public void log(int level, Object obj1,
                        Throwable exception)
        {
            if (level == FATAL)
            {
                if (log.IsFatalEnabled())
                {
                    if (obj1 != null)
                        log.Fatal(obj1, exception);
                    else
                        log.Fatal(exception);
                }
            }
            else if (level == ERROR)
            {
                if (log.IsErrorEnabled())
                {
                    if (obj1 != null)
                        log.Error(obj1, exception);
                    else
                        log.Error(exception);
                }
            }
            else if (level == WARN)
            {
                if (log.IsWarnEnabled())
                {
                    if (obj1 != null)
                        log.Warn(obj1, exception);
                    else
                        log.Warn(exception);
                }
            }
            else if (level == INFO)
            {
                if (log.IsInfoEnabled())
                {
                    if (obj1 != null)
                        log.Info(obj1, exception);
                    else
                        log.Info(exception);
                }
            }
            else if (level == DEBUG)
            {
                if (log.IsDebugEnabled())
                {
                    if (obj1 != null)
                        log.Debug(obj1, exception);
                    else
                        log.Debug(exception);
                }
            }
            else
            {
                if (log.IsTraceEnabled())
                {
                    if (obj1 != null)
                        log.Trace(obj1, exception);
                    else
                        log.Trace(exception);
                }
            }

        }

        /**
         * Check if a logger is enabled to log at the specified level
         *
         * @param level One of DEBUG, INFO, WARN, ERROR, FATAL
         */

        public bool Check(int level)
        {
            if (level == FATAL)
            {
                if (log.IsFatalEnabled())
                {
                    return true;
                }
            }
            else if (level == ERROR)
            {
                if (log.IsErrorEnabled())
                {
                    return true;
                }
            }
            else if (level == WARN)
            {
                if (log.IsWarnEnabled())
                {
                    return true;
                }
            }
            else if (level == INFO)
            {
                if (log.IsInfoEnabled())
                {
                    return true;
                }
            }
            else if (level == DEBUG)
            {
                if (log.IsDebugEnabled())
                {
                    return true;
                }
            }

            return false;

        }


    }   // end package scope class POILogger


}