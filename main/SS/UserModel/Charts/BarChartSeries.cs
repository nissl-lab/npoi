using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Charts
{
    public interface IBarChartSeries<Tx, Ty> : IChartSeries
    {
        /**
         * @return data source used for category axis data.
         */
        IChartDataSource<Tx> GetCategoryAxisData();

        /**
         * @return data source used for value axis.
         */
        IChartDataSource<Ty> GetValues();
    }
}
