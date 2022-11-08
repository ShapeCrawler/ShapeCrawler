using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

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
}