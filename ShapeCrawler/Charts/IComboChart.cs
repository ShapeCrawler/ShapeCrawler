using System.Collections.Generic;
using ShapeCrawler.Charts;
using ShapeCrawler.Collections;
using ShapeCrawler.SlideMasters;
using OneOf;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a Combination chart.
    /// </summary>
    public interface IComboChart : IChart
    {

    }

    internal class SCComboChart : SCChart, IComboChart
    {
        internal SCComboChart(P.GraphicFrame pGraphicFrame, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject)
            : base(pGraphicFrame, slideObject)
        {
        }

        public IReadOnlyList<IChart> Charts { get; }
    }
}