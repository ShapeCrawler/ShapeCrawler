using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.DevTests;

public class SeriesTests : SCTest
{
    [Test]
    public void Name_Setter_updates_existing_name_record()
    {
        // Arrange
        var pres = PresentationWithDirectSeriesName();
        var series = pres.Slide(1).Shape("Bar Chart").BarChart!.SeriesCollection[0];
        series.Name.Should().Be("Existing name");

        // Act
        series.Name = "Updated name";

        // Assert
        series.HasName.Should().BeTrue();
        series.Name.Should().Be("Updated name");

        pres = SaveAndOpenPresentation(pres);
        series = pres.Slide(1).Shape("Bar Chart").BarChart!.SeriesCollection[0];
        series.Name.Should().Be("Updated name");
        ValidatePresentation(pres);
    }

    [Test]
    public void Name_Setter_updates_linked_name_record_and_workbook()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001 bar chart.pptx"));
        var chart = pres.Slide(1).Shape("Bar Chart 1").BarChart!;
        var series = chart.SeriesCollection[0];

        // Act
        series.Name = "Updated linked name";

        // Assert
        series.Name.Should().Be("Updated linked name");
        WorksheetCellText(chart.GetWorksheetByteArray(), "B1")
            .Should().Be("Updated linked name");

        pres = SaveAndOpenPresentation(pres);
        chart = pres.Slide(1).Shape("Bar Chart 1").BarChart!;
        chart.SeriesCollection[0].Name.Should().Be("Updated linked name");
        WorksheetCellText(chart.GetWorksheetByteArray(), "B1")
            .Should().Be("Updated linked name");
        using var sdkPresentation = pres.GetSdkPresentationDocument();
        var stringReference = sdkPresentation.PresentationPart!.SlideParts
            .SelectMany(slidePart => slidePart.ChartParts)
            .Single()
            .ChartSpace!
            .Descendants<C.BarChartSeries>()
            .First()
            .SeriesText!
            .StringReference!;
        stringReference.Formula!.Text.Should().Be("Sheet1!$B$1");
        stringReference.StringCache!.PointCount!.Val!.Value.Should().Be(1U);
        stringReference.StringCache.Elements<C.StringPoint>().Single().Index!.Value
            .Should().Be(0U);
        stringReference.StringCache!.Elements<C.StringPoint>().Single().NumericValue!.Text
            .Should().Be("Updated linked name");
        ValidatePresentation(pres);
    }

    [Test]
    public void Name_Setter_does_not_update_cache_when_workbook_update_fails()
    {
        // Arrange
        using var source = TestAsset("001 bar chart.pptx");
        var stream = new MemoryStream();
        source.CopyTo(stream);
        stream.Position = 0;
        using (var sdkPresentation = PresentationDocument.Open(stream, true))
        {
            var chartPart = sdkPresentation.PresentationPart!.SlideParts
                .SelectMany(slidePart => slidePart.ChartParts)
                .Single();
            chartPart.ChartSpace!.Descendants<C.BarChartSeries>()
                .First()
                .SeriesText!
                .StringReference!
                .Formula!
                .Text = "MissingSheet!$B$1";
        }

        stream.Position = 0;
        var pres = new Presentation(stream);
        var chart = pres.Slide(1).Shape("Bar Chart 1").BarChart!;
        var series = chart.SeriesCollection[0];
        var originalName = series.Name;

        // Act
        var action = () => series.Name = "Must not commit";

        // Assert
        action.Should().Throw<InvalidOperationException>();
        series.Name.Should().Be(originalName);
        WorksheetCellText(chart.GetWorksheetByteArray(), "B1").Should().Be(originalName);
        ValidatePresentation(pres);
    }

    [Test]
    public void Name_Setter_adds_missing_name_record()
    {
        // Arrange
        var pres = PresentationWithMissingSeriesName();
        var series = pres.Slide(1).Shape("Bar Chart").BarChart!.SeriesCollection[0];
        series.HasName.Should().BeFalse();

        // Act
        series.Name = "Added name";

        // Assert
        series.HasName.Should().BeTrue();
        series.Name.Should().Be("Added name");

        pres = SaveAndOpenPresentation(pres);
        series = pres.Slide(1).Shape("Bar Chart").BarChart!.SeriesCollection[0];
        series.Name.Should().Be("Added name");
        ValidatePresentation(pres);
    }

    private static Presentation PresentationWithDirectSeriesName()
    {
        return new Presentation(p =>
        {
            p.Slide(s =>
            {
                s.ClusteredBarChartShape(chart =>
                {
                    chart.Name("Bar Chart");
                    chart.Categories("Category");
                    chart.Series("Existing name", 1);
                });
            });
        });
    }

    private static Presentation PresentationWithMissingSeriesName()
    {
        var pres = PresentationWithDirectSeriesName();
        var stream = new MemoryStream();
        pres.Save(stream);
        stream.Position = 0;
        using (var sdkPresentation = PresentationDocument.Open(stream, true))
        {
            var chartPart = sdkPresentation.PresentationPart!.SlideParts
                .SelectMany(slidePart => slidePart.ChartParts)
                .Single();
            chartPart.ChartSpace!.Descendants<C.BarChartSeries>().Single().SeriesText!.Remove();
        }

        stream.Position = 0;
        return new Presentation(stream);
    }

    private static string WorksheetCellText(byte[] workbookBytes, string address)
    {
        using var stream = new MemoryStream(workbookBytes);
        using var sdkWorkbook = SpreadsheetDocument.Open(stream, false);
        var workbookPart = sdkWorkbook.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<X.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var cell = worksheetPart.Worksheet!.Descendants<X.Cell>()
            .Single(xCell => xCell.CellReference == address);
        var value = cell.InnerText;
        return cell.DataType?.Value == X.CellValues.SharedString
            ? workbookPart.SharedStringTablePart!.SharedStringTable!
                .Elements<X.SharedStringItem>()
                .ElementAt(int.Parse(value))
                .InnerText
            : value;
    }
}
