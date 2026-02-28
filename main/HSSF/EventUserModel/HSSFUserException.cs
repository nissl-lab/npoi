/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */


namespace NPOI.HSSF.EventUserModel
{
    using System;
    /// <summary>
    /// This exception Is provided as a way for API users to throw
    /// exceptions from their event handling code. By doing so they
    /// abort file Processing by the HSSFEventFactory and by
    /// catching it from outside the HSSFEventFactory.ProcessEvents
    /// method they can diagnose the cause for the abort.
    /// The HSSFUserException supports a nested "reason"
    /// throwable, i.e. an exception that caused this one to be thrown.
    /// The HSSF package does not itself throw any of these
    /// exceptions.
    /// </summary>
    /// <remarks>
    /// @author Rainer Klute (klute@rainer-klute.de)
    /// @author Carey Sublette (careysub@earthling.net)
    /// </remarks>
    [Serializable]
    public class HSSFUserException : Exception
    {

        /// <summary>
        /// Creates a new HSSFUserException
        /// </summary>
        public HSSFUserException()
            : base()
        {

        }

        /// <summary>
        /// Creates a new HSSFUserException with a message
        /// string.
        /// </summary>
        /// <param name="msg">The MSG.</param>
        public HSSFUserException(String msg)
            : base(msg)
        {

        }

        /// <summary>
        /// Creates a new HSSFUserException with a reason.
        /// </summary>
        /// <param name="reason">The reason.</param>
        public HSSFUserException(Exception reason)
        {
            
        }

        /// <summary>
        /// Creates a new HSSFUserException with a message string
        /// and a reason.
        /// </summary>
        /// <param name="msg">The MSG.</param>
        /// <param name="reason">The reason.</param>
        public HSSFUserException(String msg, Exception reason)
            : base(msg, reason)
        {

        }
    }
}