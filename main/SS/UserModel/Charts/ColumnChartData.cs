﻿using System.Collections.Generic;

namespace NPOI.SS.UserModel.Charts
{
    /// <summary>
    /// Data for a Column Chart
    /// </summary>
    public interface IColumnChartData<Tx, Ty> : IChartData
    {
        /// <summary>
        /// Adds the series.
        /// </summary>
        /// <param name="categories">The categories data source.</param>
        /// <param name="values">The values data source.</param>
        /// <returns>Created series.</returns>
        IColumnChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categories, IChartDataSource<Ty> values);

        /// <summary>
        /// Returns list of all series.
        /// </summary>
        List<IColumnChartSeries<Tx, Ty>> GetSeries();

        /// <summary>
        /// Sets bar grouping
        /// </summary>
        /// <param name="grouping">The type of bar grouping</param>
        void SetBarGrouping(BarGrouping grouping);
    }
}
