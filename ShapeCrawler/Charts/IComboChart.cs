using System.Collections.Generic;
using ShapeCrawler.Charts;
using ShapeCrawler.Collections;
using ShapeCrawler.SlideMasters;
using OneOf;
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