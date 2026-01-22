using FluentAssertions;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class BubbleChartTests : SCTest
{
    [Test]
    public void Series_BubbleSizePoint_Value_setter_updates_bubble_size()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide(s =>
            {
                s.BubbleChartShape(shape =>
                {
                    shape.Chart(chart =>
                    {
                        chart.Name("Bubble Chart 1");
                        chart.Series(
                            "Series 1",
                            (X: 10, Y: 20, Size: 5),
                            (X: 30, Y: 40, Size: 10));
                    });
                });
            });
        });

        var bubbleChart = pres.Slide(1).Shape("Bubble Chart 1").BubbleChart!;
        var series = bubbleChart.SeriesCollection[0];

        // Act
        series.BubbleSizePoints![0].Value = 6;

        // Assert
        series.BubbleSizePoints[0].Value.Should().Be(6);
        ValidatePresentation(pres);
    }
}
