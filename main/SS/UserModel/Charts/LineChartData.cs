using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Charts
{
    /// <summary>
    /// Data for a Line Chart
    /// </summary>
    /// <typeparam name="Tx"></typeparam>
    /// <typeparam name="Ty"></typeparam>
    public interface ILineChartData<Tx, Ty> : IChartData
    {
        ILineChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categories, IChartDataSource<Ty> values);

        /**
         * @return list of all series.
         */
        List<ILineChartSeries<Tx, Ty>> GetSeries();
    }
}
