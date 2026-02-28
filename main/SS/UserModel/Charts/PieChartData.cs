using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.SS.UserModel.Charts
{
    /// <summary>
    /// Data for a Pie Chart
    /// </summary>
    public interface IPieChartData<Tx, Ty> : IChartData
    {
        /// <summary>
        /// Adds the series.
        /// </summary>
        /// <param name="categories">The categories data source.</param>
        /// <param name="values">The values data source.</param>
        /// <returns>Created series.</returns>
        IPieChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categories, IChartDataSource<Ty> values);

        /// <summary>
        /// Return list of all series.
        /// </summary>
        List<IPieChartSeries<Tx, Ty>> GetSeries();
    }
}
