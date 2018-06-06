using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Charts
{
    /// <summary>
    /// Data for a Bar Chart
    /// </summary>
    /// <typeparam name="Tx"></typeparam>
    /// <typeparam name="Ty"></typeparam>
    public interface IBarChartData<Tx, Ty> : IChartData
    {

        IBarChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categories, IChartDataSource<Ty> values);

        /**
         * @return list of all series.
         */
        List<IBarChartSeries<Tx, Ty>> GetSeries();
    }
}
