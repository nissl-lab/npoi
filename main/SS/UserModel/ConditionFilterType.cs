using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public enum ConditionFilterType
    {
        FILTER,
        TOP_10,
        UNIQUE_VALUES,
        DUPLICATE_VALUES,
        CONTAINS_TEXT,
        NOT_CONTAINS_TEXT,
        BEGINS_WITH,
        ENDS_WITH,
        CONTAINS_BLANKS,
        NOT_CONTAINS_BLANKS,
        CONTAINS_ERRORS,
        NOT_CONTAINS_ERRORS,
        TIME_PERIOD,
        ABOVE_AVERAGE
    }
}
