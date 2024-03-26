using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public interface IConditionFilterData
    {
        /// <summary>
        /// </summary>
        /// <returns>true if the flag is missing or set to true</returns>
        bool AboveAverage { get; }

        /// <summary>
        /// </summary>
        /// <returns>true if the flag is set</returns>
        bool Bottom { get; }

        /// <summary>
        /// </summary>
        /// <returns>true if the flag is set</returns>
        bool EqualAverage { get; }

        /// <summary>
        /// </summary>
        /// <returns>true if the flag is set</returns>
        bool Percent { get; }

        /// <summary>
        /// </summary>
        /// <returns>value, or 0 if not used/defined</returns>
        long Rank { get; }

        /// <summary>
        /// </summary>
        /// <returns>value, or 0 if not used/defined</returns>
        int StdDev { get; }
    }
}
