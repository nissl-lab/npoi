using System.Collections.Generic;
using NPOI.SS.UserModel.Charts;

namespace NPOI.SS.UserModel
{
    /// <summary>
    /// High level representation of a chart.
    /// </summary>
    /// @author Roman Kashitsyn
    public interface IChart : ManuallyPositionable
    {

        /// <summary>
        /// </summary>
        /// <returns>an appropriate ChartDataFactory implementation</returns>
        IChartDataFactory ChartDataFactory { get; }

        /// <summary>
        /// </summary>
        /// <returns>an appropriate ChartAxisFactory implementation</returns>
        IChartAxisFactory ChartAxisFactory { get; }

        /// <summary>
        /// </summary>
        /// <returns>chart legend instance</returns>
        IChartLegend GetOrCreateLegend();

        /// <summary>
        /// Delete current chart legend.
        /// </summary>
        void DeleteLegend();

        /// <summary>
        /// </summary>
        /// <returns>list of all chart axis</returns>
        List<IChartAxis> GetAxis();

        /// <summary>
        /// Plots specified data on the chart.
        /// </summary>
        /// <param name="data">a data to plot</param>
        void Plot(IChartData data, params IChartAxis[] axis);

        void SetTitle(string newTitle);
    }
}
