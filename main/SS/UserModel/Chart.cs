using System.Collections.Generic;
using NPOI.SS.UserModel.Charts;

namespace NPOI.SS.UserModel
{
    /**
     * High level representation of a chart.
     *
     * @author Roman Kashitsyn
     */
    public interface IChart : ManuallyPositionable
    {

        /**
         * @return an appropriate ChartDataFactory implementation
         */
        IChartDataFactory ChartDataFactory { get; }

        /**
         * @return an appropriate ChartAxisFactory implementation
         */
        IChartAxisFactory ChartAxisFactory { get; }

        /**
         * @return chart legend instance
         */
        IChartLegend GetOrCreateLegend();

        /**
         * Delete current chart legend.
         */
        void DeleteLegend();

        /**
         * @return list of all chart axis
         */
        List<IChartAxis> GetAxis();

        /**
         * Plots specified data on the chart.
         *
         * @param data a data to plot
         */
        void Plot(IChartData data, params IChartAxis[] axis);
    }
}
