using System;
using System.Collections.Generic;
using System.Text;
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
        ChartDataFactory GetChartDataFactory();

        /**
         * @return an appropriate ChartAxisFactory implementation
         */
        ChartAxisFactory GetChartAxisFactory();

        /**
         * @return chart legend instance
         */
        ChartLegend GetOrCreateLegend();

        /**
         * Delete current chart legend.
         */
        void DeleteLegend();

        /**
         * @return list of all chart axis
         */
        List<ChartAxis> GetAxis();

        /**
         * Plots specified data on the chart.
         *
         * @param data a data to plot
         */
        void Plot(ChartData data, params ChartAxis[] axis);
    }
}
