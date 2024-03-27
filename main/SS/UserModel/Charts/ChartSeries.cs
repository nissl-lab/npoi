using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Charts
{
    public enum TitleType
    {
        String,
        CellReference
    }

    public interface IChartSeries
    {
        /// <summary>
        /// Sets the title of the series as a string literal.
        /// </summary>
        /// <param name="title">title</param>
        void SetTitle(String title);

        /// <summary>
        /// Sets the title of the series as a cell reference.
        /// </summary>
        /// <param name="titleReference">titleReference</param>
        void SetTitle(CellReference titleReference);

        /// <summary>
        /// </summary>
        /// <returns>title as string literal.</returns>
        String GetTitleString();

        /// <summary>
        /// </summary>
        /// <returns>title as cell reference.</returns>
        CellReference GetTitleCellReference();

        /// <summary>
        /// </summary>
        /// <returns>title type.</returns>
        TitleType? GetTitleType();
    }
}
