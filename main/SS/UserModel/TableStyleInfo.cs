using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public interface ITableStyleInfo
    {
        bool IsShowColumnStripes { get; set; }
        /// <summary>
        /// return true if alternating row styles should be applied
        /// </summary>
        bool IsShowRowStripes { get; set; }
        /// <summary>
        /// return true if the distinct first column style should be applied
        /// </summary>
        bool IsShowFirstColumn { get; set; }

        /// <summary>
        /// return true if the distinct last column style should be applied
        /// </summary>
        bool IsShowLastColumn { get; set; }
        /// <summary>
        /// return the name of the style (may reference a built-in style)
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// style definition
        /// </summary>
        ITableStyle Style { get; }
    }
}
