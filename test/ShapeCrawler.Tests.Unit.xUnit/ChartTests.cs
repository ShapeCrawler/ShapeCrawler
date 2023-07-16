using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;
using Assert = Xunit.Assert;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
[SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class ChartTests : SCTest
{
    [Xunit.Theory]
    [MemberData(nameof(TestCasesSeriesCollectionCount))]
    public void SeriesCollection_Count_returns_number_of_series(IChart chart, int expectedSeriesCount)
    {
        // Act
        int seriesCount = chart.SeriesCollection.Count;

        // Assert
        Assert.Equal(expectedSeriesCount, seriesCount);
    }

    public static IEnumerable<object[]> TestCasesSeriesCollectionCount()
    {
        var pptxStream = GetInputStream("013.pptx");
        var presentation = SCPresentation.Open(pptxStream);
        IChart chart = (IChart) presentation.Slides[0].Shapes.First(sp => sp.Id == 5);
        yield return new object[] {chart, 3};

        pptxStream = GetInputStream("009_table.pptx");
        presentation = SCPresentation.Open(pptxStream);
        chart = (IChart) presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        yield return new object[] {chart, 1};
    }
}