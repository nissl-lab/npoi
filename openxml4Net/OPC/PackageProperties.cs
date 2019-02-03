using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.OpenXml4Net.OPC
{
    /**
     * Represents the core properties of an OPC package.
     * 
     * @author Julien Chable
     * @version 1.0
     * @see org.apache.poi.OpenXml4Net.opc.OPCPackage
     */
    public interface PackageProperties
    {
        /* Getters and Setters */

        /**
         * Set the category of the content of this package.
         */
        String GetCategoryProperty();

        /**
         * Set the category of the content of this package.
         */
        void SetCategoryProperty(String category);

        /**
         * Set the status of the content.
         */
        String GetContentStatusProperty();

        /**
         * Get the status of the content.
         */
        void SetContentStatusProperty(String contentStatus);

        /**
         * Get the type of content represented, generally defined by a specific use
         * and intended audience.
         */
        String GetContentTypeProperty();

        /**
         * Set the type of content represented, generally defined by a specific use
         * and intended audience.
         */
        void SetContentTypeProperty(String contentType);

        /**
         * Get the date of creation of the resource.
         */
        Nullable<DateTime> GetCreatedProperty();

        /**
         * Set the date of creation of the resource.
         */
        void SetCreatedProperty(String created);

        /**
         * Set the date of creation of the resource.
         */
        void SetCreatedProperty(Nullable<DateTime> created);

        /**
         * Get the entity primarily responsible for making the content of the
         * resource.
         */
        String GetCreatorProperty();

        /**
         * Set the entity primarily responsible for making the content of the
         * resource.
         */
        void SetCreatorProperty(String creator);

        /**
         * Get the explanation of the content of the resource.
         */
        String GetDescriptionProperty();

        /**
         * Set the explanation of the content of the resource.
         */
        void SetDescriptionProperty(String description);

        /**
         * Get an unambiguous reference to the resource within a given context.
         */
        String GetIdentifierProperty();

        /**
         * Set an unambiguous reference to the resource within a given context.
         */
        void SetIdentifierProperty(String identifier);

        /**
         * Get a delimited Set of keywords to support searching and indexing. This
         * is typically a list of terms that are not available elsewhere in the
         * properties
         */
        String GetKeywordsProperty();

        /**
         * Set a delimited Set of keywords to support searching and indexing. This
         * is typically a list of terms that are not available elsewhere in the
         * properties
         */
        void SetKeywordsProperty(String keywords);

        /**
         * Get the language of the intellectual content of the resource.
         */
        String GetLanguageProperty();

        /**
         * Set the language of the intellectual content of the resource.
         */
        void SetLanguageProperty(String language);

        /**
         * Get the user who performed the last modification.
         */
        String GetLastModifiedByProperty();

        /**
         * Set the user who performed the last modification.
         */
        void SetLastModifiedByProperty(String lastModifiedBy);

        /**
         * Get the date and time of the last printing.
         */
        Nullable<DateTime> GetLastPrintedProperty();

        /**
         * Set the date and time of the last printing.
         */
        void SetLastPrintedProperty(String lastPrinted);

        /**
         * Set the date and time of the last printing.
         */
        void SetLastPrintedProperty(Nullable<DateTime> lastPrinted);

        /**
         * Get the date on which the resource was changed.
         */
        Nullable<DateTime> GetModifiedProperty();

        /**
         * Set the date on which the resource was changed.
         */
        void SetModifiedProperty(String modified);

        /**
         * Set the date on which the resource was changed.
         */
        void SetModifiedProperty(Nullable<DateTime> modified);

        /**
         * Get the revision number.
         */
        String GetRevisionProperty();

        /**
         * Set the revision number.
         */
        void SetRevisionProperty(String revision);

        /**
         * Get the topic of the content of the resource.
         */
        String GetSubjectProperty();

        /**
         * Set the topic of the content of the resource.
         */
        void SetSubjectProperty(String subject);

        /**
         * Get the name given to the resource.
         */
        String GetTitleProperty();

        /**
         * Set the name given to the resource.
         */
        void SetTitleProperty(String title);

        /**
         * Get the version number.
         */
        String GetVersionProperty();

        /**
         * Set the version number.
         */
        void SetVersionProperty(String version);
    }
}
