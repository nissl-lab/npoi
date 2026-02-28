using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public interface IConditionFilterData
    {
        /**
    * @return true if the flag is missing or set to true
    */
        bool AboveAverage{ get; }

        /**
         * @return true if the flag is set
         */
        bool Bottom{ get; }

        /**
         * @return true if the flag is set
         */
        bool EqualAverage{ get; }

        /**
         * @return true if the flag is set
         */
        bool Percent{ get; }

        /**
         * @return value, or 0 if not used/defined
         */
        long Rank{ get; }

        /**
         * @return value, or 0 if not used/defined
         */
        int StdDev{ get; }
    }
}
