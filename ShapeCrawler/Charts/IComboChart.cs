using System.Collections.Generic;
using ShapeCrawler.Collections;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a Combination chart.
    /// </summary>
    public interface IComboChart : IChart
    {
        /// <summary>
        ///     Gets collection of chart category.
        /// </summary>
        public ICategoryCollection Categories { get; }
    }

    internal class SCComboChart : SCChart, IComboChart
    {
        internal SCComboChart(P.GraphicFrame pGraphicFrame, SCSlide slide)
            : base(pGraphicFrame, slide)
        {
        }

        public IReadOnlyList<IChart> Charts { get; }
    }
}