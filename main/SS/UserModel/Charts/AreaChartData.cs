using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Charts
{
    /// <summary>
    /// Data for an Area Chart
    /// </summary>
    public interface IAreaChartData<Tx, Ty> : IChartData
    {

        /// <summary>
        /// Adds the series.
        /// </summary>
        /// <param name="categories">The categories data source.</param>
        /// <param name="values">The values data source.</param>
        /// <returns>Created series.</returns>
        IAreaChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categories, IChartDataSource<Ty> values);

        /// <summary>
        /// Return list of all series.
        /// </summary>
        List<IAreaChartSeries<Tx, Ty>> GetSeries();
    }
}
