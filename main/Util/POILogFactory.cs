
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
using System.Collections;
using System.Configuration;

namespace NPOI.Util
{
    public class POILogFactory
    {

        /**
         * Map of POILogger instances, with classes as keys
         */
        private static Hashtable _loggers = new Hashtable();

        /**
         * A common instance of NullLogger, as it does nothing
         *  we only need the one
         */
        private static POILogger _nullLogger = new NullLogger();
        /**
         * The name of the class to use. Initialised the
         *  first time we need it
         */
        private static String _loggerClassName = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="POILogFactory"/> class.
        /// </summary>
        private POILogFactory()
        {
        }

        /// <summary>
        /// Get a logger, based on a class name
        /// </summary>
        /// <param name="type">the class whose name defines the log</param>
        /// <returns>a POILogger for the specified class</returns>
        public static POILogger GetLogger(Type type)
        {
            return GetLogger(type.Name);
        }
        
        /// <summary>
        /// Get a logger, based on a String
        /// </summary>
        /// <param name="cat">the String that defines the log</param>
        /// <returns>a POILogger for the specified class</returns>
        public static POILogger GetLogger(String cat)
        {
            POILogger logger = null;
            
            // If we haven't found out what logger to use yet,
            //  then do so now
            // Don't look it up until we're first asked, so
            //  that our users can set the system property
            //  between class loading and first use
            if(_loggerClassName == null) {
#if NETFRAMEWORK
        	    try {
        		    _loggerClassName = ConfigurationManager.AppSettings["loggername"];
        	    } catch(Exception) {}
#endif
                if(_loggerClassName == null)
                {
                    // try environment
                    _loggerClassName = Environment.GetEnvironmentVariable("NPOI_LOGGER_NAME");
                }

                // Use the default logger if none specified,
        	    //  or none could be fetched
        	    if(_loggerClassName == null) {
        		    _loggerClassName = _nullLogger.GetType().Name;
        	    }
            }
            
            // Short circuit for the null logger, which
            //  ignores all categories
            if(_loggerClassName.Equals(_nullLogger.GetType().Name)) {
        	    return _nullLogger;
            }

            
            // Fetch the right logger for them, creating
            //  it if that's required 
            if (_loggers.ContainsKey(cat)) {
                logger = (POILogger)_loggers[cat];
            } else {
                try {
                    //logger=assembly.CreateInstance(_loggerClassName) as POILogger;
                    Type loggerClass = Type.GetType(_loggerClassName);
                    logger =  Activator.CreateInstance(loggerClass) as POILogger;
                    logger.Initialize(cat);
                } catch(Exception) {
                  // Give up and use the null logger
                  logger = _nullLogger;
                }
                
                // Save for next time
                _loggers[cat] = logger;
            }
            return logger;
        }
    }
}
