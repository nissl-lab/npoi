
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

        private POILogger logger = null;


        public void Initialize(String cat)
        {
            this.logger = POILogFactory.GetLogger(cat);
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
                if (logger.IsFatalEnabled())
                {
                    logger.Fatal(obj1);
                }
            }
            else if (level == ERROR)
            {
                if (logger.IsErrorEnabled())
                {
                    logger.Error(obj1);
                }
            }
            else if (level == WARN)
            {
                if (logger.IsWarnEnabled())
                {
                    logger.Warn(obj1);
                }
            }
            else if (level == INFO)
            {
                if (logger.IsInfoEnabled())
                {
                    logger.Info(obj1);
                }
            }
            else if (level == DEBUG)
            {
                if (logger.IsDebugEnabled())
                {
                    logger.Debug(obj1);
                }
            }
            else
            {
                if (logger.IsTraceEnabled())
                {
                    logger.Trace(obj1);
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
                if (logger.IsFatalEnabled())
                {
                    return true;
                }
            }
            else if (level == ERROR)
            {
                if (logger.IsErrorEnabled())
                {
                    return true;
                }
            }
            else if (level == WARN)
            {
                if (logger.IsWarnEnabled())
                {
                    return true;
                }
            }
            else if (level == INFO)
            {
                if (logger.IsInfoEnabled())
                {
                    return true;
                }
            }
            else if (level == DEBUG)
            {
                if (logger.IsDebugEnabled())
                {
                    return true;
                }
            }

            return false;

        }


    }   // end package scope class POILogger


}