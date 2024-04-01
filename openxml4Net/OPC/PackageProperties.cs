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
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{
    /// <summary>
    /// Represents the core properties of an OPC package.
    /// See <see cref="OPCPackage"/>
    /// </summary>
    /// <remarks>
    /// @author Julien Chable
    /// @version 1.0
    /// </remarks>

    public interface PackageProperties
    {
        /* Getters and Setters */

        /// <summary>
        /// Set the category of the content of this package.
        /// </summary>
        String GetCategoryProperty();

        /// <summary>
        /// Set the category of the content of this package.
        /// </summary>
        void SetCategoryProperty(String category);

        /// <summary>
        /// Set the status of the content.
        /// </summary>
        String GetContentStatusProperty();

        /// <summary>
        /// Get the status of the content.
        /// </summary>
        void SetContentStatusProperty(String contentStatus);

        /// <summary>
        /// Get the type of content represented, generally defined by a specific use
        /// and intended audience.
        /// </summary>
        String GetContentTypeProperty();

        /// <summary>
        /// Set the type of content represented, generally defined by a specific use
        /// and intended audience.
        /// </summary>
        void SetContentTypeProperty(String contentType);

        /// <summary>
        /// Get the date of creation of the resource.
        /// </summary>
        Nullable<DateTime> GetCreatedProperty();

        /// <summary>
        /// Set the date of creation of the resource.
        /// </summary>
        void SetCreatedProperty(String created);

        /// <summary>
        /// Set the date of creation of the resource.
        /// </summary>
        void SetCreatedProperty(Nullable<DateTime> created);

        /// <summary>
        /// Get the entity primarily responsible for making the content of the
        /// resource.
        /// </summary>
        String GetCreatorProperty();

        /// <summary>
        /// Set the entity primarily responsible for making the content of the
        /// resource.
        /// </summary>
        void SetCreatorProperty(String creator);

        /// <summary>
        /// Get the explanation of the content of the resource.
        /// </summary>
        String GetDescriptionProperty();

        /// <summary>
        /// Set the explanation of the content of the resource.
        /// </summary>
        void SetDescriptionProperty(String description);

        /// <summary>
        /// Get an unambiguous reference to the resource within a given context.
        /// </summary>
        String GetIdentifierProperty();

        /// <summary>
        /// Set an unambiguous reference to the resource within a given context.
        /// </summary>
        void SetIdentifierProperty(String identifier);

        /// <summary>
        /// Get a delimited Set of keywords to support searching and indexing. This
        /// is typically a list of terms that are not available elsewhere in the
        /// properties
        /// </summary>
        String GetKeywordsProperty();

        /// <summary>
        /// Set a delimited Set of keywords to support searching and indexing. This
        /// is typically a list of terms that are not available elsewhere in the
        /// properties
        /// </summary>
        void SetKeywordsProperty(String keywords);

        /// <summary>
        /// Get the language of the intellectual content of the resource.
        /// </summary>
        String GetLanguageProperty();

        /// <summary>
        /// Set the language of the intellectual content of the resource.
        /// </summary>
        void SetLanguageProperty(String language);

        /// <summary>
        /// Get the user who performed the last modification.
        /// </summary>
        String GetLastModifiedByProperty();

        /// <summary>
        /// Set the user who performed the last modification.
        /// </summary>
        void SetLastModifiedByProperty(String lastModifiedBy);

        /// <summary>
        /// Get the date and time of the last printing.
        /// </summary>
        Nullable<DateTime> GetLastPrintedProperty();

        /// <summary>
        /// Set the date and time of the last printing.
        /// </summary>
        void SetLastPrintedProperty(String lastPrinted);

        /// <summary>
        /// Set the date and time of the last printing.
        /// </summary>
        void SetLastPrintedProperty(Nullable<DateTime> lastPrinted);

        /// <summary>
        /// Get the date on which the resource was changed.
        /// </summary>
        Nullable<DateTime> GetModifiedProperty();

        /// <summary>
        /// Set the date on which the resource was changed.
        /// </summary>
        void SetModifiedProperty(String modified);

        /// <summary>
        /// Set the date on which the resource was changed.
        /// </summary>
        void SetModifiedProperty(Nullable<DateTime> modified);

        /// <summary>
        /// Get the revision number.
        /// </summary>
        String GetRevisionProperty();

        /// <summary>
        /// Set the revision number.
        /// </summary>
        void SetRevisionProperty(String revision);

        /// <summary>
        /// Get the topic of the content of the resource.
        /// </summary>
        String GetSubjectProperty();

        /// <summary>
        /// Set the topic of the content of the resource.
        /// </summary>
        void SetSubjectProperty(String subject);

        /// <summary>
        /// Get the name given to the resource.
        /// </summary>
        String GetTitleProperty();

        /// <summary>
        /// Set the name given to the resource.
        /// </summary>
        void SetTitleProperty(String title);

        /// <summary>
        /// Get the version number.
        /// </summary>
        String GetVersionProperty();

        /// <summary>
        /// Set the version number.
        /// </summary>
        void SetVersionProperty(String version);
    }
}
