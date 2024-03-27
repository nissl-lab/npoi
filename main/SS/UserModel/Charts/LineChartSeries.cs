using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Charts
{
    public interface ILineChartSeries<Tx, Ty> : IChartSeries
    {
        /// <summary>
        /// </summary>
        /// <returns>data source used for category axis data.</returns>
        IChartDataSource<Tx> GetCategoryAxisData();

        /// <summary>
        /// </summary>
        /// <returns>data source used for value axis.</returns>
        IChartDataSource<Ty> GetValues();
    }
}
