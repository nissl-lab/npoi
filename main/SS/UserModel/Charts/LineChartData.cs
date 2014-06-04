using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Charts
{
    public interface ILineChartData<Tx, Ty> : IChartData
    {
        ILineChartSerie<Tx, Ty> AddSerie(IChartDataSource<Tx> categories, IChartDataSource<Ty> values);

        /**
         * @return list of all series.
         */
        List<ILineChartSerie<Tx, Ty>> GetSeries();
    }
}
