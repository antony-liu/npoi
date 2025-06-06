﻿using System.Collections.Generic;

namespace NPOI.SS.UserModel.Charts
{
    /// <summary>
    /// Data for a Bar Chart
    /// </summary>
    /// <typeparam name="Tx"></typeparam>
    /// <typeparam name="Ty"></typeparam>
    public interface IBarChartData<Tx, Ty> : IChartData
    {
        /// <summary>
        /// Adds the series.
        /// </summary>
        /// <param name="categories">The categories data source.</param>
        /// <param name="values">The values data source.</param>
        /// <returns>Created series.</returns>
        IBarChartSeries<Tx, Ty> AddSeries(IChartDataSource<Tx> categories, IChartDataSource<Ty> values);

        /**
         * @return list of all series.
         */
        List<IBarChartSeries<Tx, Ty>> GetSeries();
        
        /// <summary>
        /// Sets bar grouping
        /// </summary>
        /// <param name="grouping">The type of bar grouping</param>
        void SetBarGrouping(BarGrouping grouping);
    }
}
